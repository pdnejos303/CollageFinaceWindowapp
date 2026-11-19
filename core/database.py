import csv
import datetime
import hashlib
import os
import re
import sqlite3
from typing import Any, Dict, List, Optional, Tuple

from core.config import (
    THAI_MONTHS_LONG,
    THAI_MONTHS_SHORT,
    format_receipt_date_thai,
    parse_date_robust,
    split_receipt_prefix_and_number,
    thai_baht_text,
)
from core.security import SecurityManager

try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    import psycopg2
except ImportError:
    psycopg2 = None


def derive_year_level_and_education_level(class_group: str, default_year: int = 1, default_edu: str = "ปวช.1") -> Tuple[int, str]:
    if not class_group:
        return default_year, default_edu
    cg = str(class_group).strip()
    m = re.search(r'(ปวช|ปวส|ทล\.บ)\.?\s*(\d+)', cg)
    if m:
        prefix = m.group(1)
        yr = int(m.group(2))
        return yr, f"{prefix}.{yr}"
    m2 = re.search(r'(\d+)\s*/', cg)
    if m2:
        return int(m2.group(1)), default_edu
    m3 = re.search(r'(\d+)', cg)
    if m3:
        return int(m3.group(1)), default_edu
    return default_year, default_edu


class DatabaseManager:
    """
    High-Performance Database Manager for Suphanburi Vocational College (SPVC)
    Supports:
    - 100% Configurable Master Dropdown Options (master_options)
    - 6 Curricula & 25+ Real SPVC Fee Items (fee_catalog)
    - Single & Full-Classroom Batch Receipt Issuing (receipts + receipt_items)
    - Full Student CRUD + Atomic Excel/CSV Bulk Import
    """

    def __init__(self, db_path: str, database_url: str = ""):
        self.db_path = db_path
        self.database_url = (database_url or "").strip()
        self.engine = "postgresql" if (self.database_url.startswith("postgres") and psycopg2 is not None) else "sqlite"
        self.data_version: int = 0
        self._options_cache: Dict[str, List[str]] = {}
        os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True)
        self.initialize_schema()
        self.seed_defaults_if_empty()

    def invalidate_cache(self) -> None:
        """Clear cached options and bump data_version (e.g., after DB restore)."""
        self._options_cache.clear()
        self.data_version += 1

    @staticmethod
    def _sql_normalize_date(raw_val: Any) -> str:
        d = parse_date_robust(raw_val)
        return d.isoformat() if d else ""

    @staticmethod
    def _sql_thai_date(raw_val: Any) -> str:
        d = parse_date_robust(raw_val)
        if not d:
            return str(raw_val or "")
        byear = d.year + 543 if d.year < 2400 else d.year
        m_txt = THAI_MONTHS_SHORT[d.month] if 1 <= d.month <= 12 else ""
        m_long = THAI_MONTHS_LONG[d.month] if 1 <= d.month <= 12 else ""
        raw_s = str(raw_val or "")
        return f"{d.day}/{d.month}/{byear} {d.day:02d}/{d.month:02d}/{byear} {d.day} {m_txt} {byear} {d.day} {m_long} {byear} {d.isoformat()} {raw_s}"

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path, timeout=30.0)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON;")
        conn.execute("PRAGMA synchronous = NORMAL;")
        conn.execute("PRAGMA cache_size = -64000;")
        conn.execute("PRAGMA temp_store = MEMORY;")
        conn.create_function("sqlite_normalize_date", 1, self._sql_normalize_date)
        conn.create_function("sqlite_thai_date", 1, self._sql_thai_date)
        return conn

    def initialize_schema(self):
        with self._connect() as conn:
            conn.execute("PRAGMA journal_mode = WAL;")
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS master_options (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    category TEXT NOT NULL,
                    option_value TEXT NOT NULL,
                    sort_order INTEGER DEFAULT 10,
                    UNIQUE(category, option_value)
                );

                CREATE TABLE IF NOT EXISTS academic_terms (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    term_code TEXT UNIQUE NOT NULL,
                    term_name TEXT NOT NULL,
                    is_current INTEGER DEFAULT 0
                );

                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    full_name TEXT NOT NULL,
                    position TEXT DEFAULT 'ครู',
                    role TEXT NOT NULL DEFAULT 'ADMIN',
                    department TEXT NOT NULL DEFAULT 'งานการเงินและบัญชี',
                    active INTEGER DEFAULT 1
                );

                CREATE TABLE IF NOT EXISTS students (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    student_code TEXT UNIQUE NOT NULL,
                    full_name TEXT NOT NULL,
                    citizen_id TEXT NOT NULL DEFAULT '',
                    education_level TEXT NOT NULL DEFAULT 'ปวช.1',
                    class_group TEXT NOT NULL DEFAULT 'ปวช.1/1',
                    faculty TEXT NOT NULL DEFAULT 'พาณิชยกรรม',
                    major TEXT NOT NULL DEFAULT 'การบัญชี',
                    year_level INTEGER NOT NULL DEFAULT 1,
                    scholarship_type TEXT DEFAULT 'ปกติ (Regular)',
                    registered_credits INTEGER NOT NULL DEFAULT 20,
                    total_fee REAL NOT NULL DEFAULT 0.0,
                    paid_amount REAL NOT NULL DEFAULT 0.0,
                    status TEXT NOT NULL DEFAULT 'UNPAID',
                    phone TEXT DEFAULT ''
                );

                CREATE TABLE IF NOT EXISTS fee_catalog (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    fee_code TEXT UNIQUE NOT NULL,
                    curriculum_name TEXT NOT NULL DEFAULT 'หลักสูตรในระบบ-ปวช',
                    fee_name TEXT NOT NULL,
                    calc_type TEXT NOT NULL DEFAULT 'FIXED',
                    default_credits REAL NOT NULL DEFAULT 0.0,
                    default_price REAL NOT NULL DEFAULT 0.0,
                    education_level TEXT NOT NULL DEFAULT 'ALL',
                    is_waivable_15yr INTEGER DEFAULT 0,
                    sort_order INTEGER DEFAULT 10,
                    is_active INTEGER DEFAULT 1
                );

                CREATE TABLE IF NOT EXISTS student_enrollments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    student_id INTEGER NOT NULL,
                    term_code TEXT NOT NULL,
                    registered_credits INTEGER NOT NULL DEFAULT 20,
                    total_fee REAL NOT NULL DEFAULT 0.0,
                    discount_amount REAL NOT NULL DEFAULT 0.0,
                    net_fee REAL NOT NULL DEFAULT 0.0,
                    paid_amount REAL NOT NULL DEFAULT 0.0,
                    balance_amount REAL NOT NULL DEFAULT 0.0,
                    payment_status TEXT NOT NULL DEFAULT 'UNPAID',
                    note TEXT DEFAULT '',
                    UNIQUE(student_id, term_code),
                    FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS receipts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    receipt_no TEXT UNIQUE NOT NULL,
                    book_no TEXT NOT NULL DEFAULT '7006',
                    receipt_date TEXT NOT NULL,
                    term_code TEXT NOT NULL DEFAULT '1/2569',
                    student_id INTEGER,
                    student_code TEXT NOT NULL,
                    snapshot_student_name TEXT NOT NULL,
                    snapshot_education_level TEXT NOT NULL DEFAULT 'ปวช.1',
                    snapshot_class_group TEXT NOT NULL DEFAULT '',
                    snapshot_department TEXT NOT NULL DEFAULT '',
                    snapshot_major TEXT NOT NULL DEFAULT '',
                    subtotal_amount REAL NOT NULL DEFAULT 0.0,
                    discount_amount REAL NOT NULL DEFAULT 0.0,
                    net_total_amount REAL NOT NULL DEFAULT 0.0,
                    amount_text_th TEXT NOT NULL DEFAULT '',
                    payment_method TEXT NOT NULL DEFAULT 'เงินสด (Cash)',
                    payment_ref TEXT DEFAULT '',
                    evidence_path TEXT DEFAULT '',
                    cashier_username TEXT NOT NULL DEFAULT 'admin',
                    cashier_name TEXT NOT NULL DEFAULT 'เจ้าหน้าที่การเงิน',
                    cashier_position TEXT NOT NULL DEFAULT 'ครู',
                    status TEXT NOT NULL DEFAULT 'ISSUED',
                    void_reason TEXT DEFAULT '',
                    voided_at TEXT DEFAULT '',
                    voided_by TEXT DEFAULT '',
                    printed_count INTEGER DEFAULT 1,
                    remark TEXT DEFAULT ''
                );

                CREATE TABLE IF NOT EXISTS receipt_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    receipt_id INTEGER NOT NULL,
                    line_no INTEGER NOT NULL DEFAULT 1,
                    fee_code TEXT NOT NULL DEFAULT 'CUSTOM',
                    curriculum_name TEXT NOT NULL DEFAULT '',
                    item_name TEXT NOT NULL,
                    quantity REAL NOT NULL DEFAULT 1.0,
                    unit_price REAL NOT NULL DEFAULT 0.0,
                    discount REAL NOT NULL DEFAULT 0.0,
                    line_total REAL NOT NULL DEFAULT 0.0,
                    FOREIGN KEY(receipt_id) REFERENCES receipts(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS audit_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TEXT NOT NULL,
                    actor TEXT NOT NULL,
                    role TEXT NOT NULL,
                    action TEXT NOT NULL,
                    target TEXT NOT NULL,
                    detail TEXT NOT NULL,
                    prev_hash TEXT NOT NULL,
                    curr_hash TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS course_catalog (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    course_code TEXT UNIQUE NOT NULL,
                    course_name TEXT NOT NULL,
                    theory_hours INTEGER NOT NULL DEFAULT 0,
                    practice_hours INTEGER NOT NULL DEFAULT 2,
                    credits INTEGER NOT NULL DEFAULT 1,
                    major TEXT NOT NULL DEFAULT 'เทคโนโลยีธุรกิจดิจิทัล',
                    class_level TEXT NOT NULL DEFAULT 'ปวช.1',
                    sort_order INTEGER DEFAULT 10
                );

                CREATE TABLE IF NOT EXISTS student_course_registrations (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    student_code TEXT NOT NULL,
                    term_code TEXT NOT NULL DEFAULT '2/2569',
                    line_no INTEGER NOT NULL DEFAULT 1,
                    course_code TEXT NOT NULL,
                    course_name TEXT NOT NULL,
                    theory_hours INTEGER NOT NULL DEFAULT 0,
                    practice_hours INTEGER NOT NULL DEFAULT 2,
                    credits INTEGER NOT NULL DEFAULT 1
                );

                CREATE INDEX IF NOT EXISTS idx_master_options_cat ON master_options(category, sort_order);
                CREATE INDEX IF NOT EXISTS idx_students_code_name ON students(student_code, full_name, class_group, status);
                CREATE INDEX IF NOT EXISTS idx_students_dept_level ON students(education_level, faculty, major);
                CREATE INDEX IF NOT EXISTS idx_students_status_fee ON students(status, total_fee, paid_amount);
                CREATE INDEX IF NOT EXISTS idx_students_class_code ON students(class_group, student_code);
                CREATE INDEX IF NOT EXISTS idx_students_major_code ON students(major, student_code);
                CREATE INDEX IF NOT EXISTS idx_receipts_no_date ON receipts(receipt_no, receipt_date, status);
                CREATE INDEX IF NOT EXISTS idx_receipts_status_amt ON receipts(status, net_total_amount);
                CREATE INDEX IF NOT EXISTS idx_receipts_date_id ON receipts(receipt_date, id DESC);
                CREATE INDEX IF NOT EXISTS idx_receipts_num_no ON receipts(CAST(receipt_no AS INTEGER)) WHERE receipt_no GLOB '[0-9]*';
                CREATE INDEX IF NOT EXISTS idx_receipts_student ON receipts(student_code, term_code);
                CREATE INDEX IF NOT EXISTS idx_receipt_items_rcp ON receipt_items(receipt_id, line_no);
                CREATE INDEX IF NOT EXISTS idx_stu_course_reg ON student_course_registrations(student_code, term_code);
                """
            )

            # Safe schema migrations for existing SQLite files
            fee_cols = [r["name"] for r in conn.execute("PRAGMA table_info(fee_catalog)").fetchall()]
            if "curriculum_name" not in fee_cols:
                conn.execute("ALTER TABLE fee_catalog ADD COLUMN curriculum_name TEXT DEFAULT 'หลักสูตรในระบบ-ปวช'")
            if "default_credits" not in fee_cols:
                conn.execute("ALTER TABLE fee_catalog ADD COLUMN default_credits REAL DEFAULT 0.0")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_fee_catalog_curr ON fee_catalog(curriculum_name, is_active)")

            rcp_cols = [r["name"] for r in conn.execute("PRAGMA table_info(receipts)").fetchall()]
            if "cashier_position" not in rcp_cols:
                conn.execute("ALTER TABLE receipts ADD COLUMN cashier_position TEXT DEFAULT 'ครู'")

            item_cols = [r["name"] for r in conn.execute("PRAGMA table_info(receipt_items)").fetchall()]
            if "curriculum_name" not in item_cols:
                conn.execute("ALTER TABLE receipt_items ADD COLUMN curriculum_name TEXT DEFAULT ''")

            conn.commit()

    # ------------------------------------------------------------------
    # Master Dropdown Options (Configurable in Settings)
    # ------------------------------------------------------------------
    def get_options(self, category: str) -> List[str]:
        cat = (category or "").strip().upper()
        cached = self._options_cache.get(cat)
        if cached is not None:
            return list(cached)
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT option_value FROM master_options WHERE category = ? ORDER BY sort_order ASC, id ASC",
                (cat,),
            ).fetchall()
            opts = [r["option_value"] for r in rows]

            if cat == "TERM":
                import datetime
                base_year = datetime.datetime.now().year + 543
                standard_terms = []
                for yr in range(base_year - 2, base_year + 5):
                    for sem in ("1", "2", "ฤดูร้อน"):
                        standard_terms.append(f"{sem}/{yr}")
                for t in standard_terms:
                    if t not in opts:
                        opts.append(t)

                def _term_sort_key(t: str):
                    parts = t.split("/")
                    if len(parts) == 2 and parts[1].isdigit():
                        yr = int(parts[1])
                        sem_map = {"1": 1, "2": 2, "ฤดูร้อน": 3}
                        sem = sem_map.get(parts[0], 9)
                        return (yr, sem, t)
                    return (9999, 9, t)

                opts.sort(key=_term_sort_key)

            self._options_cache[cat] = list(opts)
            return list(opts)

    def add_option(self, category: str, value: str, actor: str = "admin", role: str = "ADMIN") -> bool:
        val = (value or "").strip()
        cat = (category or "").strip().upper()
        if not cat or not val:
            raise ValueError("กรุณาระบุข้อมูลตัวเลือกที่ต้องการเพิ่ม")
        with self._connect() as conn:
            conn.execute(
                "INSERT OR IGNORE INTO master_options (category, option_value) VALUES (?, ?)",
                (cat, val),
            )
            conn.commit()
        self._options_cache.pop(cat, None)
        self.data_version += 1
        self.log_audit(actor, role, "OPTION_ADD", f"{cat}:{val}", f"Added dropdown option '{val}' to {cat}")
        return True

    def update_option(
        self, category: str, old_value: str, new_value: str, actor: str = "admin", role: str = "ADMIN"
    ) -> bool:
        old_val = (old_value or "").strip()
        new_val = (new_value or "").strip()
        cat = (category or "").strip().upper()
        if not cat or not new_val:
            raise ValueError("กรุณาระบุข้อมูลตัวเลือกใหม่")
        with self._connect() as conn:
            cur = conn.execute(
                "UPDATE master_options SET option_value = ? WHERE category = ? AND option_value = ?",
                (new_val, cat, old_val),
            )
            if cur.rowcount == 0:
                conn.execute(
                    "INSERT OR IGNORE INTO master_options (category, option_value) VALUES (?, ?)",
                    (cat, new_val),
                )
            conn.commit()
        self._options_cache.pop(cat, None)
        self.data_version += 1
        self.log_audit(actor, role, "OPTION_UPDATE", f"{cat}:{new_val}", f"Updated option from '{old_val}' to '{new_val}' in {cat}")
        return True

    def delete_option(self, category: str, value: str, actor: str = "admin", role: str = "ADMIN") -> bool:
        val = (value or "").strip()
        cat = (category or "").strip().upper()
        with self._connect() as conn:
            conn.execute(
                "DELETE FROM master_options WHERE category = ? AND option_value = ?",
                (cat, val),
            )
            conn.commit()
        self._options_cache.pop(cat, None)
        self.data_version += 1
        self.log_audit(actor, role, "OPTION_DELETE", f"{cat}:{val}", f"Removed dropdown option '{val}' from {cat}")
        return True

    # ------------------------------------------------------------------
    # Audit Log
    # ------------------------------------------------------------------
    def log_audit(self, actor: str, role: str, action: str, target: str, detail: str):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self._connect() as conn:
            row = conn.execute("SELECT curr_hash FROM audit_logs ORDER BY id DESC LIMIT 1").fetchone()
            prev_hash = row["curr_hash"] if row else "GENESIS_0000000000000000"
            raw_payload = f"{timestamp}|{actor}|{role}|{action}|{target}|{detail}|{prev_hash}"
            curr_hash = hashlib.sha256(raw_payload.encode("utf-8")).hexdigest()[:24].upper()
            conn.execute(
                """
                INSERT INTO audit_logs (timestamp, actor, role, action, target, detail, prev_hash, curr_hash)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (timestamp, actor, role, action, target, detail, prev_hash, curr_hash),
            )
            conn.commit()

    def verify_audit_integrity(self) -> Tuple[bool, int, str]:
        with self._connect() as conn:
            rows = conn.execute("SELECT * FROM audit_logs ORDER BY id ASC").fetchall()
        expected_prev = "GENESIS_0000000000000000"
        for r in rows:
            if r["prev_hash"] != expected_prev:
                return False, len(rows), f"Chain broken at Audit ID #{r['id']}"
            raw = f"{r['timestamp']}|{r['actor']}|{r['role']}|{r['action']}|{r['target']}|{r['detail']}|{r['prev_hash']}"
            calc_hash = hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24].upper()
            if calc_hash != r["curr_hash"]:
                return False, len(rows), f"Hash mismatch at Audit ID #{r['id']}"
            expected_prev = r["curr_hash"]
        return True, len(rows), "OK (ฐานข้อมูลสมบูรณ์ 100%)"

    def get_audit_logs(
        self,
        limit: int = 200,
        search: str = "",
        action_filter: str = "ALL",
        date_from: str = "",
        date_to: str = "",
        offset: int = 0,
    ) -> List[Dict[str, Any]]:
        where_clauses: List[str] = []
        params: List[Any] = []

        act = (action_filter or "").strip()
        if act and act.upper() != "ALL" and act != "ทั้งหมด":
            where_clauses.append("action = ?")
            params.append(act)

        df = (date_from or "").strip()
        if df:
            where_clauses.append("substr(timestamp, 1, 10) >= ?")
            params.append(df)

        dt = (date_to or "").strip()
        if dt:
            where_clauses.append("substr(timestamp, 1, 10) <= ?")
            params.append(dt)

        q = (search or "").strip().lower()
        if q:
            where_clauses.append(
                "(INSTR(LOWER(COALESCE(actor, '')), ?) > 0 OR "
                "INSTR(LOWER(COALESCE(role, '')), ?) > 0 OR "
                "INSTR(LOWER(COALESCE(action, '')), ?) > 0 OR "
                "INSTR(LOWER(COALESCE(target, '')), ?) > 0 OR "
                "INSTR(LOWER(COALESCE(detail, '')), ?) > 0)"
            )
            params.extend([q, q, q, q, q])

        sql = "SELECT * FROM audit_logs"
        if where_clauses:
            sql += " WHERE " + " AND ".join(where_clauses)
        sql += " ORDER BY id DESC LIMIT ? OFFSET ?"
        params.extend([max(1, int(limit)), max(0, int(offset))])

        with self._connect() as conn:
            rows = conn.execute(sql, tuple(params)).fetchall()
            return [dict(r) for r in rows]

    def get_audit_action_types(self) -> List[str]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT DISTINCT action FROM audit_logs WHERE action IS NOT NULL AND action != '' ORDER BY action ASC"
            ).fetchall()
            return [r["action"] for r in rows]

    # ------------------------------------------------------------------
    # Authentication
    # ------------------------------------------------------------------
    def authenticate(self, username: str, password: str) -> Optional[Dict[str, Any]]:
        u = (username or "").strip()
        if not u or not password:
            return None
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM users WHERE username = ? AND active = 1", (u,)
            ).fetchone()
            if not row:
                return None
            if SecurityManager.verify_password(password, row["password_hash"]):
                return dict(row)
            return None

    # ------------------------------------------------------------------
    # Default Seed Data Matching EX Screenshots (วิทยาลัยอาชีวศึกษาสุพรรณบุรี)
    # ------------------------------------------------------------------
    def seed_defaults_if_empty(self):
        with self._connect() as conn:
            opt_cnt = conn.execute("SELECT COUNT(*) as c FROM master_options").fetchone()["c"]
            if opt_cnt == 0:
                default_options = [
                    # 6 หลักสูตรตามภาพ EX (91698_0.jpg)
                    ("CURRICULUM", "หลักสูตรในระบบ-ปวช", 1),
                    ("CURRICULUM", "หลักสูตรในระบบ-ปวส", 2),
                    ("CURRICULUM", "หลักสูตรปริญญาตรี", 3),
                    ("CURRICULUM", "ค่าวัสดุฝึกสำเร็จรูป", 4),
                    ("CURRICULUM", "หลักสูตรระยะสั้น", 5),
                    ("CURRICULUM", "อื่นๆ", 6),
                    # คณะ / ประเภทวิชา
                    ("FACULTY", "บริหารธุรกิจ / พาณิชยกรรม", 1),
                    ("FACULTY", "คหกรรมศาสตร์", 2),
                    ("FACULTY", "ศิลปกรรม", 3),
                    ("FACULTY", "อุตสาหกรรมท่องเที่ยวและการโรงแรม", 4),
                    ("FACULTY", "เทคโนโลยีสารสนเทศและการสื่อสาร", 5),
                    # แผนกวิชา / สาขาวิชา
                    ("DEPARTMENT", "การบัญชี", 1),
                    ("DEPARTMENT", "การตลาด", 2),
                    ("DEPARTMENT", "เทคโนโลยีธุรกิจดิจิทัล / คอมพิวเตอร์ธุรกิจ", 3),
                    ("DEPARTMENT", "การเลขานุการ / การจัดการสำนักงาน", 4),
                    ("DEPARTMENT", "ธุรกิจค้าปลีก", 5),
                    ("DEPARTMENT", "อาหารและโภชนาการ", 6),
                    ("DEPARTMENT", "คหกรรมศาสตร์ทั่วไป", 7),
                    ("DEPARTMENT", "แฟชั่นและสิ่งทอ", 8),
                    ("DEPARTMENT", "คอมพิวเตอร์กราฟิก", 9),
                    ("DEPARTMENT", "วิจิตรศิลป์ / การออกแบบ", 10),
                    ("DEPARTMENT", "การโรงแรม", 11),
                    ("DEPARTMENT", "การท่องเที่ยว", 12),
                    ("DEPARTMENT", "เทคโนโลยีสารสนเทศ", 13),
                    # ชั้นปี / ห้องเรียน
                    ("CLASS_LEVEL", "ปวช.1/1", 1),
                    ("CLASS_LEVEL", "ปวช.1/2", 2),
                    ("CLASS_LEVEL", "ปวช.2/1", 3),
                    ("CLASS_LEVEL", "ปวช.2/2", 4),
                    ("CLASS_LEVEL", "ปวช.3/1", 5),
                    ("CLASS_LEVEL", "ปวส.1/1", 6),
                    ("CLASS_LEVEL", "ปวส.1/2 (ทวิภาคี)", 7),
                    ("CLASS_LEVEL", "ปวส.2/1", 8),
                    ("CLASS_LEVEL", "ปวส.2/2 (ทวิภาคี)", 9),
                    ("CLASS_LEVEL", "ทล.บ.1/1", 10),
                    ("CLASS_LEVEL", "ทล.บ.2/1", 11),
                    # ภาค/ปีการศึกษา
                    ("TERM", "1/2569", 1),
                    ("TERM", "2/2569", 2),
                    ("TERM", "ฤดูร้อน/2569", 3),
                    ("TERM", "1/2570", 4),
                    # ตำแหน่งผู้รับเงิน
                    ("CASHIER_POSITION", "ครู", 1),
                    ("CASHIER_POSITION", "เจ้าหน้าที่การเงิน", 2),
                    ("CASHIER_POSITION", "หัวหน้างานการเงินและบัญชี", 3),
                ]
                conn.executemany(
                    "INSERT OR IGNORE INTO master_options (category, option_value, sort_order) VALUES (?, ?, ?)",
                    default_options,
                )

            user_count = conn.execute("SELECT COUNT(*) as c FROM users").fetchone()["c"]
            if user_count == 0:
                users = [
                    ("admin", SecurityManager.hash_password("admin123"), "นางสาวสุพรรณี ศรีอาชีวะ", "ครู", "ADMIN", "งานการเงิน วิทยาลัยอาชีวศึกษาสุพรรณบุรี"),
                    ("finance", SecurityManager.hash_password("fin123"), "คุณกนกวรรณ รับชำระค่าเทอม", "เจ้าหน้าที่การเงิน", "FINANCE", "งานการเงิน วิทยาลัยอาชีวศึกษาสุพรรณบุรี"),
                ]
                conn.executemany(
                    "INSERT INTO users (username, password_hash, full_name, position, role, department) VALUES (?, ?, ?, ?, ?, ?)",
                    users,
                )

            # Seed the 25+ Real SPVC Fee Items from EX (91700_0.jpg)
            fee_count = conn.execute("SELECT COUNT(*) as c FROM fee_catalog").fetchone()["c"]
            if fee_count == 0:
                real_spvc_fees = [
                    # (fee_code, curriculum_name, fee_name, calc_type, default_credits, default_price)
                    ("SPVC-01", "หลักสูตรในระบบ-ปวช", "ค่าคู่มือนักเรียนนักศึกษาใหม่", "FIXED", 0, 100.0),
                    ("SPVC-02", "หลักสูตรในระบบ-ปวช", "ค่าวารสารประชาสัมพันธ์", "FIXED", 0, 20.0),
                    ("SPVC-03", "หลักสูตรในระบบ-ปวช", "ค่าบำรุงห้องสมุด", "FIXED", 0, 50.0),
                    ("SPVC-04", "หลักสูตรในระบบ-ปวช", "ค่าบำรุงสุขภาพหรือห้องพยาบาลหรือค่าตรวจโรค", "FIXED", 0, 50.0),
                    ("SPVC-05", "หลักสูตรในระบบ-ปวช", "เงินรายได้สถานศึกษา", "FIXED", 0, 300.0),
                    ("SPVC-06", "หลักสูตรในระบบ-ปวช", "ค่ารักษาสภาพแวดล้อม", "FIXED", 0, 200.0),
                    ("SPVC-07", "หลักสูตรในระบบ-ปวช", "ค่าใช้บริการอินเตอร์เน็ต", "FIXED", 0, 300.0),
                    ("SPVC-08", "หลักสูตรในระบบ-ปวช", "ค่าจ้างครูหรือผู้ที่มีความเชี่ยวชาญในสาขาเฉพาะ", "FIXED", 0, 500.0),
                    ("SPVC-09", "หลักสูตรในระบบ-ปวช", "ค่ากีฬานันทนาการ", "FIXED", 0, 100.0),
                    ("SPVC-10", "หลักสูตรในระบบ-ปวช", "ค่าเพิ่มมาตรฐานคุณภาพของสถานศึกษา", "FIXED", 0, 200.0),
                    ("SPVC-11", "หลักสูตรในระบบ-ปวช", "ค่าประกันอุบัติเหตุนักเรียน/นักศึกษา", "FIXED", 0, 250.0),
                    # ปวส.
                    ("SPVC-12", "หลักสูตรในระบบ-ปวส", "ค่าลงทะเบียนรายวิชาปฏิบัติ (หน่วยกิตละ 100)", "PER_CREDIT", 20, 100.0),
                    ("SPVC-13", "หลักสูตรในระบบ-ปวส", "ค่าลงทะเบียนรายวิชาทฤษฎี (หน่วยกิตละ 50)", "PER_CREDIT", 18, 50.0),
                    ("SPVC-14", "หลักสูตรในระบบ-ปวส", "ค่ารักษาสภาพแวดล้อม (ปวส.)", "FIXED", 0, 330.0),
                    ("SPVC-15", "หลักสูตรในระบบ-ปวส", "ค่าใช้บริการอินเตอร์เน็ต (ปวส.)", "FIXED", 0, 300.0),
                    ("SPVC-16", "หลักสูตรในระบบ-ปวส", "เงินรายได้สถานศึกษา (ปวส.)", "FIXED", 0, 300.0),
                    ("SPVC-17", "หลักสูตรในระบบ-ปวส", "ค่าจ้างครูหรือผู้ที่มีความเชี่ยวชาญในสาขาเฉพาะ (ปวส.)", "FIXED", 0, 500.0),
                    ("SPVC-18", "หลักสูตรในระบบ-ปวส", "ค่าบำรุงห้องสมุด (ปวส.)", "FIXED", 0, 50.0),
                    ("SPVC-19", "หลักสูตรในระบบ-ปวส", "ค่าบำรุงสุขภาพหรือห้องพยาบาล (ปวส.)", "FIXED", 0, 50.0),
                    # ปริญญาตรี (ทล.บ.)
                    ("SPVC-20", "หลักสูตรปริญญาตรี", "ค่าหน่วยกิตระดับปริญญาตรี (ทล.บ.)", "PER_CREDIT", 18, 300.0),
                    ("SPVC-21", "หลักสูตรปริญญาตรี", "ค่าธรรมเนียมบำรุงการศึกษาปริญญาตรี", "FIXED", 0, 2500.0),
                    # ค่าวัสดุฝึกสำเร็จรูป
                    ("SPVC-22", "ค่าวัสดุฝึกสำเร็จรูป", "ค่าวัสดุฝึกหรืออุปกรณ์การศึกษา", "FIXED", 0, 20.0),
                    ("SPVC-23", "ค่าวัสดุฝึกสำเร็จรูป", "ค่าชุดฝึกปฏิบัติการวิชาชีพสำเร็จรูป", "FIXED", 0, 450.0),
                    # หลักสูตรระยะสั้น
                    ("SPVC-24", "หลักสูตรระยะสั้น", "ค่าลงทะเบียนเรียนหลักสูตรวิชาชีพระยะสั้น", "FIXED", 0, 300.0),
                    # อื่นๆ (ตามภาพ 91700_0.jpg)
                    ("SPVC-25", "อื่นๆ", "ค่าลงหน่วยกิตรายวิชาทฤษฎีภาคฤดูร้อน", "PER_CREDIT", 3, 250.0),
                    ("SPVC-26", "อื่นๆ", "ค่าลงทะเบียนเรียนรายวิชาปฏิบัติภาคฤดูร้อน", "PER_CREDIT", 3, 300.0),
                    ("SPVC-27", "อื่นๆ", "ค่าปรับในกรณีที่มีการลงทะเบียนรายวิชาช้ากว่ากำหนด", "FIXED", 0, 50.0),
                    ("SPVC-28", "อื่นๆ", "ค่าลงทะเบียนโอนรายวิชาเรียนภาคฤดูร้อน", "FIXED", 0, 150.0),
                    ("SPVC-29", "อื่นๆ", "เรียนซ้ำภาคฤดูร้อน", "FIXED", 0, 250.0),
                    ("SPVC-30", "อื่นๆ", "ค่าลงทะเบียนรายวิชาปรับพื้นฐาน", "FIXED", 0, 50.0),
                    ("SPVC-31", "อื่นๆ", "ค่าสอบเทียบความรู้หรือประสบการณ์", "FIXED", 0, 10.0),
                    ("SPVC-32", "อื่นๆ", "ค่าลงทะเบียนเพื่อรักษาสภาพการเป็นนักเรียน", "FIXED", 0, 100.0),
                    ("SPVC-33", "อื่นๆ", "ค่าคืนสภาพการเป็นนักเรียน", "FIXED", 0, 100.0),
                    ("SPVC-34", "อื่นๆ", "ค่าออกสำเนาระเบียนแสดงผลการเรียน", "FIXED", 0, 20.0),
                    ("SPVC-35", "อื่นๆ", "ค่าคู่มือรับสมัครศึกษาต่อ", "FIXED", 0, 50.0),
                ]
                for idx, (code, curr, name, ctype, cred, price) in enumerate(real_spvc_fees, start=1):
                    conn.execute(
                        """
                        INSERT INTO fee_catalog
                        (fee_code, curriculum_name, fee_name, calc_type, default_credits, default_price, sort_order, is_active)
                        VALUES (?, ?, ?, ?, ?, ?, ?, 1)
                        """,
                        (code, curr, name, ctype, cred, price, idx),
                    )

            stu_count = conn.execute("SELECT COUNT(*) as c FROM students").fetchone()["c"]
            if stu_count == 0:
                students = [
                    ("692020101", "นายธนภัทร ศรีสุพรรณ", "1729900123456", "ปวช.1", "ปวช.1/1", "บริหารธุรกิจ / พาณิชยกรรม", "การบัญชี", 1, "เรียนฟรี 15 ปี (ปวช.)", 18, 950.0, 950.0, "PAID", "081-234-5678"),
                    ("692020102", "นางสาวพิมชนก วงศ์สุวรรณ", "1729900234567", "ปวช.1", "ปวช.1/1", "บริหารธุรกิจ / พาณิชยกรรม", "การบัญชี", 1, "เรียนฟรี 15 ปี (ปวช.)", 18, 950.0, 0.0, "UNPAID", "089-876-5432"),
                    ("692020103", "นางสาวกมลชนก สุพรรณภูมิ", "1729900234599", "ปวช.1", "ปวช.1/1", "บริหารธุรกิจ / พาณิชยกรรม", "การบัญชี", 1, "เรียนฟรี 15 ปี (ปวช.)", 18, 950.0, 0.0, "UNPAID", "089-111-2233"),
                    ("692090101", "นายกิตติพงศ์ เจริญยิ่ง", "1729900345678", "ปวช.2", "ปวช.2/1", "ศิลปกรรม", "คอมพิวเตอร์กราฟิก", 2, "ปกติ (Regular)", 20, 1570.0, 0.0, "UNPAID", "082-345-6789"),
                    ("693020101", "นางสาวชนากานต์ อัศวเดช", "1729900456789", "ปวส.1", "ปวส.1/1", "บริหารธุรกิจ / พาณิชยกรรม", "การบัญชี", 1, "ปกติ (Regular)", 20, 3230.0, 3230.0, "PAID", "086-112-3344"),
                    ("693020201", "นายณัฐวุฒิ ทองคำแท้", "1729900567890", "ปวส.1", "ปวส.1/2 (ทวิภาคี)", "บริหารธุรกิจ / พาณิชยกรรม", "การตลาด", 1, "ทวิภาคี (Dual Vocational)", 19, 2850.0, 0.0, "UNPAID", "084-556-7788"),
                    ("693040101", "นางสาวอริสรา ใจสว่าง", "1729900678901", "ปวส.2", "ปวส.2/1", "คหกรรมศาสตร์", "อาหารและโภชนาการ", 2, "กยศ. (Student Loan)", 21, 3450.0, 0.0, "UNPAID", "085-667-8899"),
                ]
                conn.executemany(
                    """
                    INSERT INTO students
                    (student_code, full_name, citizen_id, education_level, class_group, faculty, major, year_level, scholarship_type, registered_credits, total_fee, paid_amount, status, phone)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    students,
                )

            rcp_count = conn.execute("SELECT COUNT(*) as c FROM receipts").fetchone()["c"]
            if rcp_count == 0:
                today_iso = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                sample_rcp = [
                    (
                        "700605433", "7006", today_iso, "1/2569", 1, "692020101",
                        "นายธนภัทร ศรีสุพรรณ", "ปวช.1", "ปวช.1/1", "บริหารธุรกิจ / พาณิชยกรรม", "การบัญชี",
                        950.0, 0.0, 950.0, thai_baht_text(950.0),
                        "เงินสด (Cash)", "", "", "admin", "นางสาวสุพรรณี ศรีอาชีวะ", "ครู", "ISSUED", "",
                        [
                            (1, "SPVC-01", "หลักสูตรในระบบ-ปวช", "ค่าคู่มือนักเรียนนักศึกษาใหม่", 0, 100.0, 0.0, 100.0),
                            (2, "SPVC-03", "หลักสูตรในระบบ-ปวช", "ค่าบำรุงห้องสมุด", 0, 50.0, 0.0, 50.0),
                            (3, "SPVC-05", "หลักสูตรในระบบ-ปวช", "เงินรายได้สถานศึกษา", 0, 300.0, 0.0, 300.0),
                            (4, "SPVC-06", "หลักสูตรในระบบ-ปวช", "ค่ารักษาสภาพแวดล้อม", 0, 200.0, 0.0, 200.0),
                            (5, "SPVC-07", "หลักสูตรในระบบ-ปวช", "ค่าใช้บริการอินเตอร์เน็ต", 0, 300.0, 0.0, 300.0),
                        ],
                    ),
                ]
                for (
                    r_no, b_no, r_date, t_code, s_id, s_code, s_name, s_lvl, s_cls, s_fac, s_maj,
                    sub_amt, disc_amt, net_amt, txt_th, p_method, p_ref, ev_p, c_user, c_name, c_pos, st, rem, items
                ) in sample_rcp:
                    cur = conn.execute(
                        """
                        INSERT INTO receipts
                        (receipt_no, book_no, receipt_date, term_code, student_id, student_code,
                         snapshot_student_name, snapshot_education_level, snapshot_class_group, snapshot_department, snapshot_major,
                         subtotal_amount, discount_amount, net_total_amount, amount_text_th,
                         payment_method, payment_ref, evidence_path, cashier_username, cashier_name, cashier_position, status, remark)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            r_no, b_no, r_date, t_code, s_id, s_code, s_name, s_lvl, s_cls, s_fac, s_maj,
                            sub_amt, disc_amt, net_amt, txt_th, p_method, p_ref, ev_p, c_user, c_name, c_pos, st, rem
                        ),
                    )
                    rcp_id = cur.lastrowid
                    for ln, f_code, curr_n, i_name, qty, u_price, disc, l_tot in items:
                        conn.execute(
                            """
                            INSERT INTO receipt_items
                            (receipt_id, line_no, fee_code, curriculum_name, item_name, quantity, unit_price, discount, line_total)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (rcp_id, ln, f_code, curr_n, i_name, qty, u_price, disc, l_tot),
                        )

            # Ensure EXofBill reference data (69219100049 / Receipt 700605308) exists
            conn.executemany(
                "INSERT OR IGNORE INTO master_options (category, option_value, sort_order) VALUES (?, ?, ?)",
                [
                    ("FACULTY", "พาณิชยกรรม/บริหารธุรกิจ", 1),
                    ("DEPARTMENT", "เทคโนโลยีธุรกิจดิจิทัล", 2),
                    ("CLASS_LEVEL", "ปวช.1", 0),
                ],
            )
            ex_stu = conn.execute("SELECT id FROM students WHERE student_code = '69219100049'").fetchone()
            if not ex_stu:
                cur_s = conn.execute(
                    """
                    INSERT INTO students
                    (student_code, full_name, citizen_id, education_level, class_group, faculty, major, year_level, scholarship_type, registered_credits, total_fee, paid_amount, status, phone)
                    VALUES ('69219100049', 'นางสาวปรันดา เจียศิริพันธ์', '1729900999999', 'ปวช.1', 'ปวช.1', 'พาณิชยกรรม/บริหารธุรกิจ', 'เทคโนโลยีธุรกิจดิจิทัล', 1, 'เรียนฟรี 15 ปี (ปวช.)', 19, 1130.0, 1130.0, 'PAID', '')
                    """
                )
                ex_stu_id = cur_s.lastrowid
            else:
                ex_stu_id = ex_stu["id"]

            ex_rcp = conn.execute("SELECT id FROM receipts WHERE receipt_no = '700605308'").fetchone()
            if not ex_rcp:
                cur_r = conn.execute(
                    """
                    INSERT INTO receipts
                    (receipt_no, book_no, receipt_date, term_code, student_id, student_code,
                     snapshot_student_name, snapshot_education_level, snapshot_class_group, snapshot_department, snapshot_major,
                     subtotal_amount, discount_amount, net_total_amount, amount_text_th,
                     payment_method, payment_ref, evidence_path, cashier_username, cashier_name, cashier_position, status, remark)
                    VALUES ('700605308', '7006', '2026-10-30', '2/2569', ?, '69219100049',
                            'นางสาวปรันดา เจียศิริพันธ์', 'ปวช.1', 'ปวช.1', 'พาณิชยกรรม/บริหารธุรกิจ', 'เทคโนโลยีธุรกิจดิจิทัล',
                            1130.0, 0.0, 1130.0, ?, 'เงินสด (Cash)', '', '', 'admin', 'นางสาวสุพรรณี ศรีอาชีวะ', 'ครู', 'ISSUED', '')
                    """,
                    (ex_stu_id, thai_baht_text(1130.0)),
                )
                ex_rcp_id = cur_r.lastrowid
                ex_items = [
                    (1, "SPVC-06", "หลักสูตรในระบบ-ปวช", "ค่ารักษาสภาพแวดล้อม", 1.0, 330.0, 0.0, 330.0),
                    (2, "SPVC-07", "หลักสูตรในระบบ-ปวช", "ค่าใช้บริการอินเตอร์เน็ต", 1.0, 300.0, 0.0, 300.0),
                    (3, "SPVC-08", "หลักสูตรในระบบ-ปวช", "ค่าจ้างครูหรือผู้ที่มีความเชี่ยวชาญในสาขาเฉพาะ", 1.0, 500.0, 0.0, 500.0),
                ]
                conn.executemany(
                    """
                    INSERT INTO receipt_items
                    (receipt_id, line_no, fee_code, curriculum_name, item_name, quantity, unit_price, discount, line_total)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    [(ex_rcp_id, *it) for it in ex_items],
                )

            # Seed the 11 Official Courses from EXofBill (อศท.28)
            ex_courses = [
                ("20000-1102", "ภาษาไทยเพื่ออาชีพ", 0, 2, 1, "เทคโนโลยีธุรกิจดิจิทัล", "ปวช.1", 1),
                ("20000-1203", "การฟังและการพูดภาษาอังกฤษ", 0, 2, 1, "เทคโนโลยีธุรกิจดิจิทัล", "ปวช.1", 2),
                ("20000-1222", "ภาษาจีนเพื่อการสื่อสารในชีวิตประจำวัน", 0, 2, 1, "เทคโนโลยีธุรกิจดิจิทัล", "ปวช.1", 3),
                ("20000-1501", "หน้าที่พลเมืองและศีลธรรม", 2, 0, 2, "เทคโนโลยีธุรกิจดิจิทัล", "ปวช.1", 4),
                ("20001-1003", "ธุรกิจเบื้องต้น", 1, 2, 2, "เทคโนโลยีธุรกิจดิจิทัล", "ปวช.1", 5),
                ("20001-1005", "การใช้เทคโนโลยีดิจิทัลเพื่ออาชีพ", 2, 2, 3, "เทคโนโลยีธุรกิจดิจิทัล", "ปวช.1", 6),
                ("21910-2004", "โปรแกรมประมวลผลคำ", 1, 2, 2, "เทคโนโลยีธุรกิจดิจิทัล", "ปวช.1", 7),
                ("21910-2005", "โปรแกรมตารางคำนวณ", 1, 2, 2, "เทคโนโลยีธุรกิจดิจิทัล", "ปวช.1", 8),
                ("21910-2006", "โปรแกรมนำเสนอ", 2, 2, 3, "เทคโนโลยีธุรกิจดิจิทัล", "ปวช.1", 9),
                ("21910-2009", "คณิตศาสตร์คอมพิวเตอร์", 1, 2, 2, "เทคโนโลยีธุรกิจดิจิทัล", "ปวช.1", 10),
                ("20000-2002", "กิจกรรมลูกเสือวิสามัญ 2", 0, 2, 0, "เทคโนโลยีธุรกิจดิจิทัล", "ปวช.1", 11),
            ]
            conn.executemany(
                """
                INSERT OR IGNORE INTO course_catalog
                (course_code, course_name, theory_hours, practice_hours, credits, major, class_level, sort_order)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                ex_courses,
            )

            reg_cnt = conn.execute(
                "SELECT COUNT(*) as c FROM student_course_registrations WHERE student_code = '69219100049'"
            ).fetchone()["c"]
            if reg_cnt == 0:
                for ln_idx, (ccode, cname, th, pr, cr, _mj, _cl, _so) in enumerate(ex_courses, start=1):
                    conn.execute(
                        """
                        INSERT INTO student_course_registrations
                        (student_code, term_code, line_no, course_code, course_name, theory_hours, practice_hours, credits)
                        VALUES ('69219100049', '2/2569', ?, ?, ?, ?, ?, ?)
                        """,
                        (ln_idx, ccode, cname, th, pr, cr),
                    )
            conn.commit()

        with self._connect() as conn:
            a_count = conn.execute("SELECT COUNT(*) as c FROM audit_logs").fetchone()["c"]
        if a_count == 0:
            self.log_audit("SYSTEM", "SYSTEM", "SYSTEM_INIT", "DATABASE", "Initialized SPVC Smart Receipt Database.")

    # ------------------------------------------------------------------
    # Fee Catalog Management (Grouped by 6 Curricula from EX)
    # ------------------------------------------------------------------
    def get_fee_catalog(self, curriculum_filter: str = "ALL", active_only: bool = True) -> List[Dict[str, Any]]:
        query = "SELECT * FROM fee_catalog WHERE 1=1"
        params: List[Any] = []
        if active_only:
            query += " AND is_active = 1"
        if curriculum_filter and curriculum_filter not in ("ALL", "ทุกหลักสูตร"):
            query += " AND curriculum_name = ?"
            params.append(curriculum_filter.strip())
        query += " ORDER BY sort_order ASC, id ASC"
        with self._connect() as conn:
            rows = conn.execute(query, params).fetchall()
            return [dict(r) for r in rows]

    def upsert_fee_item(
        self,
        fee_code: str,
        fee_name: str,
        calc_type: str = "FIXED",
        default_price: float = 0.0,
        is_waivable_15yr: int = 0,
        actor: str = "admin",
        role: str = "ADMIN",
        curriculum_name: str = "หลักสูตรในระบบ-ปวช",
        default_credits: float = 0.0,
    ) -> bool:
        name_clean = (fee_name or "").strip()
        if not name_clean:
            raise ValueError("กรุณาระบุชื่อรายการค่าธรรมเนียม")
        code_clean = (fee_code or "").strip().upper()
        with self._connect() as conn:
            if not code_clean:
                cnt = conn.execute("SELECT COUNT(*) as c FROM fee_catalog").fetchone()["c"]
                code_clean = f"SPVC-{cnt + 1:02d}"

            existing = conn.execute("SELECT id FROM fee_catalog WHERE fee_code = ?", (code_clean,)).fetchone()
            if existing:
                conn.execute(
                    """
                    UPDATE fee_catalog
                    SET curriculum_name = ?, fee_name = ?, calc_type = ?, default_credits = ?, default_price = ?, is_waivable_15yr = ?, is_active = 1
                    WHERE fee_code = ?
                    """,
                    (curriculum_name.strip(), name_clean, calc_type, float(default_credits), float(default_price), int(is_waivable_15yr), code_clean),
                )
            else:
                conn.execute(
                    """
                    INSERT INTO fee_catalog
                    (fee_code, curriculum_name, fee_name, calc_type, default_credits, default_price, is_waivable_15yr, is_active)
                    VALUES (?, ?, ?, ?, ?, ?, ?, 1)
                    """,
                    (code_clean, curriculum_name.strip(), name_clean, calc_type, float(default_credits), float(default_price), int(is_waivable_15yr)),
                )
            conn.commit()
        self.data_version += 1
        self.log_audit(actor, role, "FEE_UPSERT", code_clean, f"[{curriculum_name}] {name_clean} = {default_price:,.2f} THB")
        return True

    def delete_fee_item(self, fee_code: str, actor: str = "admin", role: str = "ADMIN") -> bool:
        with self._connect() as conn:
            conn.execute("DELETE FROM fee_catalog WHERE fee_code = ?", (fee_code.strip(),))
            conn.commit()
        self.data_version += 1
        self.log_audit(actor, role, "FEE_DELETE", fee_code, f"Deleted fee item {fee_code}")
        return True

    def build_default_receipt_items_for_student(self, student: Dict[str, Any]) -> List[Dict[str, Any]]:
        cls_grp = str(student.get("class_group", "ปวช.1/1"))
        if "ปวส" in cls_grp:
            return [
                {"fee_code": "SPVC-14", "curriculum_name": "หลักสูตรในระบบ-ปวส", "item_name": "ค่ารักษาสภาพแวดล้อม (ปวส.)", "quantity": 0.0, "unit_price": 330.0, "discount": 0.0, "line_total": 330.0},
                {"fee_code": "SPVC-15", "curriculum_name": "หลักสูตรในระบบ-ปวส", "item_name": "ค่าใช้บริการอินเตอร์เน็ต (ปวส.)", "quantity": 0.0, "unit_price": 300.0, "discount": 0.0, "line_total": 300.0},
                {"fee_code": "SPVC-16", "curriculum_name": "หลักสูตรในระบบ-ปวส", "item_name": "เงินรายได้สถานศึกษา (ปวส.)", "quantity": 0.0, "unit_price": 300.0, "discount": 0.0, "line_total": 300.0},
            ]
        return [
            {"fee_code": "SPVC-03", "curriculum_name": "หลักสูตรในระบบ-ปวช", "item_name": "ค่าบำรุงห้องสมุด", "quantity": 0.0, "unit_price": 50.0, "discount": 0.0, "line_total": 50.0},
            {"fee_code": "SPVC-05", "curriculum_name": "หลักสูตรในระบบ-ปวช", "item_name": "เงินรายได้สถานศึกษา", "quantity": 0.0, "unit_price": 300.0, "discount": 0.0, "line_total": 300.0},
            {"fee_code": "SPVC-06", "curriculum_name": "หลักสูตรในระบบ-ปวช", "item_name": "ค่ารักษาสภาพแวดล้อม", "quantity": 0.0, "unit_price": 200.0, "discount": 0.0, "line_total": 200.0},
            {"fee_code": "SPVC-07", "curriculum_name": "หลักสูตรในระบบ-ปวช", "item_name": "ค่าใช้บริการอินเตอร์เน็ต", "quantity": 0.0, "unit_price": 300.0, "discount": 0.0, "line_total": 300.0},
            {"fee_code": "SPVC-09", "curriculum_name": "หลักสูตรในระบบ-ปวช", "item_name": "ค่ากีฬานันทนาการ", "quantity": 0.0, "unit_price": 100.0, "discount": 0.0, "line_total": 100.0},
        ]

    # ------------------------------------------------------------------
    # Receipt Numbering (Supports custom letters/prefix & auto-incrementing numbers)
    # ------------------------------------------------------------------
    def get_next_receipt_no(self, prefix: str = "", start_no: str = "700605433", end_no: str = "700999999") -> str:
        """
        Determines the next sequential receipt number based on prefix, starting number,
        and existing receipts stored in the database.

        If prefix is supplied (e.g. 'RC-', 'INV-', 'บิล-') or present in start_no:
            - Extracts prefix and numeric portion.
            - Scans existing receipts matching this prefix.
            - Finds the highest trailing number used so far.
            - Increments by 1 (preserving padding length).
        If no prefix (pure numeric, e.g. '700605433'):
            - Scans MAX(CAST(receipt_no AS INTEGER)) for numeric receipts.
            - Increments by 1 (preserving padding length).
        """
        pfx_in = (prefix or "").strip()
        start_in = (start_no or "700605433").strip()

        parsed_pfx, parsed_num = split_receipt_prefix_and_number(start_in)

        if pfx_in:
            eff_prefix = pfx_in
            eff_num = parsed_num if parsed_num else (start_in if start_in.isdigit() else "1")
        elif parsed_pfx:
            eff_prefix = parsed_pfx
            eff_num = parsed_num or "1"
        else:
            eff_prefix = ""
            eff_num = start_in or "1"

        start_int = int(eff_num) if eff_num.isdigit() else 1
        pad_len = len(eff_num) if eff_num.isdigit() else 5
        max_seen = start_int - 1

        with self._connect() as conn:
            if eff_prefix:
                rows = conn.execute(
                    "SELECT receipt_no FROM receipts WHERE receipt_no LIKE ? ORDER BY length(receipt_no) DESC, receipt_no DESC LIMIT 2000",
                    (f"{eff_prefix}%",),
                ).fetchall()
                for r in rows:
                    r_no = str(r["receipt_no"]).strip()
                    suffix = r_no[len(eff_prefix):]
                    m = re.match(r"^(\d+)", suffix)
                    if m:
                        val = int(m.group(1))
                        if val > max_seen:
                            max_seen = val
            else:
                max_row = conn.execute(
                    "SELECT MAX(CAST(receipt_no AS INTEGER)) as max_no FROM receipts WHERE receipt_no GLOB '[0-9]*'"
                ).fetchone()
                if max_row and max_row["max_no"] is not None:
                    max_seen = max(max_seen, int(max_row["max_no"]))

        next_int = max(start_int, max_seen + 1)
        next_num_str = f"{next_int:0{pad_len}d}"
        return f"{eff_prefix}{next_num_str}"

    def _normalize_line_items(self, line_items: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], float, float, float]:
        if not line_items:
            raise ValueError("กรุณาเพิ่มรายการค่าธรรมเนียมในใบเสร็จอย่างน้อย 1 รายการ")

        normalized_items = []
        subtotal = 0.0
        total_discount = 0.0
        net_total = 0.0

        for idx, item in enumerate(line_items, start=1):
            raw_qty = item.get("quantity")
            raw_price = item.get("unit_price")
            raw_disc = item.get("discount")
            try:
                qty = round(float(raw_qty if raw_qty is not None else 0.0), 2)
            except (ValueError, TypeError):
                qty = 0.0
            try:
                u_price = round(float(raw_price if raw_price is not None else 0.0), 2)
            except (ValueError, TypeError):
                u_price = 0.0
            try:
                disc = round(float(raw_disc if raw_disc is not None else 0.0), 2)
            except (ValueError, TypeError):
                disc = 0.0

            eff_qty = qty if qty > 0 else 1.0
            gross_line = round(eff_qty * u_price, 2)
            line_tot = max(0.0, round(gross_line - disc, 2))
            subtotal += gross_line
            total_discount += disc
            net_total += line_tot
            normalized_items.append(
                {
                    "line_no": idx,
                    "fee_code": str(item.get("fee_code") or "CUSTOM").strip() or "CUSTOM",
                    "curriculum_name": str(item.get("curriculum_name") or "").strip(),
                    "item_name": str(item.get("item_name") or "ค่าธรรมเนียมการศึกษา").strip(),
                    "quantity": qty,
                    "unit_price": u_price,
                    "discount": disc,
                    "line_total": line_tot,
                }
            )

        subtotal = round(subtotal, 2)
        total_discount = round(total_discount, 2)
        net_total = round(net_total, 2)
        if net_total < 0:
            raise ValueError("ยอดรวมเป็นเงินในใบเสร็จต้องไม่ต่ำกว่า 0.00 บาท")

        return normalized_items, subtotal, total_discount, net_total

    def issue_flexible_receipt(
        self,
        student_code: str,
        line_items: List[Dict[str, Any]],
        payment_method: str,
        actor: str,
        cashier_name: str,
        role: str = "FINANCE",
        term_code: str = "1/2569",
        book_no: str = "7006",
        receipt_prefix: str = "",
        receipt_no_override: str = "",
        receipt_date_override: str = "",
        cashier_position: str = "ครู",
        payment_ref: str = "",
        evidence_path: str = "",
        remark: str = "",
        update_tuition_balance: bool = True,
        custom_payer_name: str = "",
        custom_faculty: str = "",
        custom_major: str = "",
        custom_class_group: str = "",
        start_no: str = "700605433",
        end_no: str = "700999999",
    ) -> Dict[str, Any]:
        normalized_items, subtotal, total_discount, net_total = self._normalize_line_items(line_items)

        baht_str = thai_baht_text(net_total)
        d_override = (receipt_date_override or "").strip()
        now_str = d_override if d_override else datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        r_override = (receipt_no_override or "").strip()
        receipt_no = r_override if r_override else self.get_next_receipt_no(receipt_prefix, start_no, end_no)

        _, r_num = split_receipt_prefix_and_number(receipt_no)
        _, e_num = split_receipt_prefix_and_number(end_no)
        if r_num.isdigit() and e_num.isdigit() and int(e_num) > 0:
            if int(r_num) > int(e_num):
                raise ValueError(f"เลขที่ใบเสร็จ {receipt_no} เกินช่วงเลขที่ใบเสร็จสิ้นสุด ({end_no}) ที่ตั้งค่าไว้ กรุณาขยายช่วงเลขใบเสร็จในหน้าตั้งค่า")

        s_code_in = (student_code or "").strip()
        custom_payer = (custom_payer_name or "").strip()
        custom_faculty_clean = (custom_faculty or "").strip()
        custom_major_clean = (custom_major or "").strip()
        custom_class_clean = (custom_class_group or "").strip()
        pref_clean = (payment_ref or "").strip()
        rem_clean = (remark or "").strip()

        with self._connect() as conn:
            stu = conn.execute("SELECT * FROM students WHERE student_code = ?", (s_code_in,)).fetchone() if s_code_in else None
            if stu:
                s_id = stu["id"]
                s_code = stu["student_code"]
                s_name = custom_payer or stu["full_name"]
                s_lvl = stu["education_level"]
                s_cls = custom_class_clean or stu["class_group"]
                s_dept = custom_faculty_clean or stu["faculty"]
                s_maj = custom_major_clean or stu["major"]
                old_paid = float(stu["paid_amount"] or 0.0)
                total_fee = float(stu["total_fee"] or 0.0)
            else:
                if not custom_payer:
                    raise ValueError(f"ไม่พบรหัสนักศึกษา '{student_code}' (หากเป็นผู้ชำระทั่วไป กรุณากรอกช่อง 'ได้รับเงินจาก')")
                s_id = None
                s_code = s_code_in or "EXTERNAL"
                s_name = custom_payer
                s_lvl = custom_class_clean or "ปวช.1"
                s_cls = custom_class_clean or "ปวช.1/1"
                s_dept = custom_faculty_clean or "บริหารธุรกิจ / พาณิชยกรรม"
                s_maj = custom_major_clean or "การบัญชี"
                old_paid = 0.0
                total_fee = net_total

            cur = conn.execute(
                """
                INSERT INTO receipts
                (receipt_no, book_no, receipt_date, term_code, student_id, student_code,
                 snapshot_student_name, snapshot_education_level, snapshot_class_group, snapshot_department, snapshot_major,
                 subtotal_amount, discount_amount, net_total_amount, amount_text_th,
                 payment_method, payment_ref, evidence_path, cashier_username, cashier_name, cashier_position, status, remark)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'ISSUED', ?)
                """,
                (
                    receipt_no,
                    book_no,
                    now_str,
                    term_code,
                    s_id,
                    s_code,
                    s_name,
                    s_lvl,
                    s_cls,
                    s_dept,
                    s_maj,
                    subtotal,
                    total_discount,
                    net_total,
                    baht_str,
                    payment_method,
                    pref_clean,
                    evidence_path or "",
                    actor,
                    cashier_name,
                    cashier_position,
                    rem_clean,
                ),
            )
            receipt_id = cur.lastrowid

            conn.executemany(
                """
                INSERT INTO receipt_items
                (receipt_id, line_no, fee_code, curriculum_name, item_name, quantity, unit_price, discount, line_total)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    (
                        receipt_id,
                        it["line_no"],
                        it["fee_code"],
                        it["curriculum_name"],
                        it["item_name"],
                        it["quantity"],
                        it["unit_price"],
                        it["discount"],
                        it["line_total"],
                    )
                    for it in normalized_items
                ],
            )

            new_paid = old_paid
            new_status = "PAID"
            new_balance = 0.0

            if stu and update_tuition_balance:
                new_paid = round(old_paid + net_total, 2)
                if new_paid >= total_fee:
                    new_status = "PAID"
                elif new_paid > 0:
                    new_status = "PARTIAL"
                else:
                    new_status = "UNPAID"
                new_balance = max(0.0, round(total_fee - new_paid, 2))
                conn.execute(
                    "UPDATE students SET paid_amount = ?, status = ? WHERE id = ?",
                    (new_paid, new_status, s_id),
                )

            conn.commit()

        self.data_version += 1
        self.log_audit(
            actor,
            role,
            "RECEIPT_ISSUED",
            receipt_no,
            f"Issued receipt {receipt_no} for {s_name} ({s_code}): {net_total:,.2f} THB",
        )

        return {
            "receipt_id": receipt_id,
            "receipt_no": receipt_no,
            "tx_code": receipt_no,
            "book_no": book_no,
            "receipt_date": now_str,
            "tx_date": now_str,
            "term_code": term_code,
            "student_code": s_code,
            "full_name": s_name,
            "snapshot_student_name": s_name,
            "education_level": s_lvl,
            "snapshot_education_level": s_lvl,
            "class_group": s_cls,
            "snapshot_class_group": s_cls,
            "faculty": s_dept,
            "snapshot_department": s_dept,
            "major": s_maj,
            "snapshot_major": s_maj,
            "subtotal_amount": subtotal,
            "discount_amount": total_discount,
            "net_total_amount": net_total,
            "amount": net_total,
            "amount_text_th": baht_str,
            "total_fee": total_fee,
            "paid_amount": new_paid,
            "balance": new_balance,
            "payment_method": payment_method,
            "payment_ref": payment_ref,
            "status": new_status,
            "receipt_status": "ISSUED",
            "recorded_by": cashier_name,
            "cashier_position": cashier_position,
            "evidence_path": evidence_path or "",
            "remark": remark,
            "items": normalized_items,
        }

    def issue_batch_class_receipts(
        self,
        students_list: List[Dict[str, Any]],
        line_items: List[Dict[str, Any]],
        payment_method: str,
        actor: str,
        cashier_name: str,
        role: str = "ADMIN",
        term_code: str = "1/2569",
        book_no: str = "7006",
        receipt_prefix: str = "",
        cashier_position: str = "ครู",
        start_no: str = "700605433",
        end_no: str = "700999999",
        receipt_date_override: str = "",
    ) -> List[Dict[str, Any]]:
        """Issues identical receipt items for a whole classroom/group of students inside a single atomic transaction."""
        if not students_list:
            raise ValueError("ไม่มีรายชื่อนักเรียน/นักศึกษาในกลุ่มที่เลือก")
        normalized_items, subtotal, total_discount, net_total = self._normalize_line_items(line_items)
        baht_str = thai_baht_text(net_total)
        d_override = (receipt_date_override or "").strip()
        now_str = d_override if d_override else datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        first_no_str = self.get_next_receipt_no(receipt_prefix, start_no, end_no)
        b_prefix, b_num_str = split_receipt_prefix_and_number(first_no_str)
        pad_len = len(b_num_str) if b_num_str else 5
        start_int = int(b_num_str) if (b_num_str and b_num_str.isdigit()) else 1

        _, e_num_str = split_receipt_prefix_and_number(end_no)
        end_int = int(e_num_str) if (e_num_str and e_num_str.isdigit() and int(e_num_str) > 0) else 0

        if end_int > 0 and (start_int + len(students_list) - 1) > end_int:
            raise ValueError(
                f"เลขที่ใบเสร็จจะเกินช่วงสิ้นสุด ({end_no}) ที่ตั้งค่าไว้ กรุณาขยายช่วงเลขใบเสร็จในหน้าตั้งค่า"
            )

        issued_receipts: List[Dict[str, Any]] = []
        with self._connect() as conn:
            conn.execute("BEGIN TRANSACTION;")
            try:
                for idx, stu_in in enumerate(students_list):
                    curr_num = start_int + idx
                    receipt_no = f"{b_prefix}{curr_num:0{pad_len}d}"

                    s_code_in = str(stu_in["student_code"]).strip()
                    stu = conn.execute("SELECT * FROM students WHERE student_code = ?", (s_code_in,)).fetchone()
                    if stu:
                        s_id = stu["id"]
                        s_code = stu["student_code"]
                        s_name = str(stu_in.get("full_name", "")).strip() or stu["full_name"]
                        s_lvl = stu["education_level"]
                        s_cls = str(stu_in.get("class_group", "")).strip() or stu["class_group"]
                        s_dept = str(stu_in.get("faculty", "")).strip() or stu["faculty"]
                        s_maj = str(stu_in.get("major", "")).strip() or stu["major"]
                        old_paid = float(stu["paid_amount"])
                        total_fee = float(stu["total_fee"])
                    else:
                        s_id = None
                        s_code = s_code_in or "EXTERNAL"
                        s_name = str(stu_in.get("full_name", "")).strip() or s_code
                        s_lvl = str(stu_in.get("class_group", "")).strip() or "ปวช.1"
                        s_cls = str(stu_in.get("class_group", "")).strip() or "ปวช.1/1"
                        s_dept = str(stu_in.get("faculty", "")).strip() or "บริหารธุรกิจ / พาณิชยกรรม"
                        s_maj = str(stu_in.get("major", "")).strip() or "การบัญชี"
                        old_paid = 0.0
                        total_fee = net_total

                    # Dynamic per-student line items support ("เผื่อบางคนไม่เหมือนเพื่อน")
                    stu_custom_items = stu_in.get("line_items") or stu_in.get("custom_items")
                    if stu_custom_items:
                        stu_norm_items, s_subtotal, s_discount, s_net = self._normalize_line_items(stu_custom_items)
                    elif "custom_net_total" in stu_in:
                        s_net = round(float(stu_in["custom_net_total"]), 2)
                        s_subtotal = s_net
                        s_discount = 0.0
                        custom_title = stu_in.get("custom_remark") or "ชำระค่าธรรมเนียมการศึกษา (กำหนดเฉพาะบุคคล)"
                        stu_norm_items = [
                            {
                                "line_no": 1,
                                "fee_code": "CUSTOM_BATCH",
                                "curriculum_name": str(stu_in.get("class_group", "")),
                                "item_name": custom_title,
                                "quantity": 1.0,
                                "unit_price": s_net,
                                "discount": 0.0,
                                "line_total": s_net,
                            }
                        ]
                    else:
                        stu_norm_items = normalized_items
                        s_subtotal = subtotal
                        s_discount = total_discount
                        s_net = net_total

                    s_baht_str = thai_baht_text(s_net)

                    cur = conn.execute(
                        """
                        INSERT INTO receipts
                        (receipt_no, book_no, receipt_date, term_code, student_id, student_code,
                         snapshot_student_name, snapshot_education_level, snapshot_class_group, snapshot_department, snapshot_major,
                         subtotal_amount, discount_amount, net_total_amount, amount_text_th,
                         payment_method, payment_ref, evidence_path, cashier_username, cashier_name, cashier_position, status, remark)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '', '', ?, ?, ?, 'ISSUED', ?)
                        """,
                        (
                            receipt_no,
                            book_no,
                            now_str,
                            term_code,
                            s_id,
                            s_code,
                            s_name,
                            s_lvl,
                            s_cls,
                            s_dept,
                            s_maj,
                            s_subtotal,
                            s_discount,
                            s_net,
                            s_baht_str,
                            payment_method,
                            actor,
                            cashier_name,
                            cashier_position,
                            stu_in.get("remark", ""),
                        ),
                    )
                    receipt_id = cur.lastrowid

                    conn.executemany(
                        """
                        INSERT INTO receipt_items
                        (receipt_id, line_no, fee_code, curriculum_name, item_name, quantity, unit_price, discount, line_total)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        [
                            (
                                receipt_id,
                                it["line_no"],
                                it["fee_code"],
                                it["curriculum_name"],
                                it["item_name"],
                                it["quantity"],
                                it["unit_price"],
                                it["discount"],
                                it["line_total"],
                            )
                            for it in stu_norm_items
                        ],
                    )

                    new_paid = old_paid
                    new_status = "PAID"
                    new_balance = 0.0
                    if stu:
                        new_paid = round(old_paid + s_net, 2)
                        if new_paid >= total_fee:
                            new_status = "PAID"
                        elif new_paid > 0:
                            new_status = "PARTIAL"
                        else:
                            new_status = "UNPAID"
                        new_balance = max(0.0, round(total_fee - new_paid, 2))
                        conn.execute(
                            "UPDATE students SET paid_amount = ?, status = ? WHERE id = ?",
                            (new_paid, new_status, s_id),
                        )

                    issued_receipts.append(
                        {
                            "receipt_id": receipt_id,
                            "receipt_no": receipt_no,
                            "tx_code": receipt_no,
                            "book_no": book_no,
                            "receipt_date": now_str,
                            "tx_date": now_str,
                            "term_code": term_code,
                            "student_code": s_code,
                            "full_name": s_name,
                            "snapshot_student_name": s_name,
                            "education_level": s_lvl,
                            "snapshot_education_level": s_lvl,
                            "class_group": s_cls,
                            "snapshot_class_group": s_cls,
                            "faculty": s_dept,
                            "snapshot_department": s_dept,
                            "major": s_maj,
                            "snapshot_major": s_maj,
                            "subtotal_amount": s_subtotal,
                            "discount_amount": s_discount,
                            "net_total_amount": s_net,
                            "amount": s_net,
                            "amount_text_th": s_baht_str,
                            "total_fee": total_fee,
                            "paid_amount": new_paid,
                            "balance": new_balance,
                            "payment_method": payment_method,
                            "payment_ref": "",
                            "status": new_status,
                            "receipt_status": "ISSUED",
                            "recorded_by": cashier_name,
                            "cashier_position": cashier_position,
                            "evidence_path": "",
                            "remark": stu_in.get("remark", ""),
                            "items": [dict(it) for it in stu_norm_items],
                        }
                    )
                conn.commit()
            except Exception:
                conn.rollback()
                raise

        if issued_receipts:
            self.data_version += 1
            self.log_audit(
                actor,
                role,
                "RECEIPT_BATCH_ISSUED",
                f"{issued_receipts[0]['receipt_no']}..{issued_receipts[-1]['receipt_no']}",
                f"Batch issued {len(issued_receipts)} receipts ({net_total:,.2f} THB each).",
            )
        return issued_receipts

    def record_tuition_payment(
        self,
        student_code: str,
        pay_amount: float,
        payment_method: str,
        actor: str,
        role: str,
        note: str = "",
        evidence_path: str = "",
    ) -> Dict[str, Any]:
        if pay_amount <= 0:
            raise ValueError("จำนวนเงินที่ชำระต้องมากกว่า 0 บาท")
        line_items = [
            {
                "fee_code": "SPVC-05",
                "item_name": note.strip() if note.strip() else "ชำระค่าธรรมเนียมการศึกษารายภาคเรียน",
                "quantity": 1.0,
                "unit_price": float(pay_amount),
                "discount": 0.0,
            }
        ]
        return self.issue_flexible_receipt(
            student_code=student_code,
            line_items=line_items,
            payment_method=payment_method,
            actor=actor,
            cashier_name=actor,
            role=role,
            evidence_path=evidence_path,
            remark=note,
        )

    # ------------------------------------------------------------------
    # Receipts Register, Date Filtering (like EX 91694_0.jpg), & Voiding
    # ------------------------------------------------------------------
    def get_distinct_receipt_dates(self) -> List[str]:
        """Returns distinct ISO date strings (YYYY-MM-DD) from all existing receipts in descending order."""
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT DISTINCT sqlite_normalize_date(receipt_date) as d FROM receipts "
                "WHERE sqlite_normalize_date(receipt_date) != '' ORDER BY d DESC LIMIT 365"
            ).fetchall()
            return [r["d"] for r in rows if r["d"]]

    def get_distinct_receipt_dates_formatted(self) -> List[Dict[str, str]]:
        """Returns formatted distinct receipt dates suitable for UI dropdowns with Thai Buddhist formatting."""
        iso_dates = self.get_distinct_receipt_dates()
        today = datetime.date.today()
        today_iso = today.isoformat()
        if today_iso not in iso_dates:
            iso_dates.insert(0, today_iso)

        res: List[Dict[str, str]] = []
        for iso_d in iso_dates:
            d_obj = parse_date_robust(iso_d)
            if not d_obj:
                continue
            byear = d_obj.year + 543 if d_obj.year < 2400 else d_obj.year
            m_txt = THAI_MONTHS_SHORT[d_obj.month] if 1 <= d_obj.month <= 12 else ""
            thai_short = f"{d_obj.day:02d}/{d_obj.month:02d}/{byear}"
            if d_obj == today:
                label = f"{thai_short} (วันนี้)"
            elif d_obj == today - datetime.timedelta(days=1):
                label = f"{thai_short} (เมื่อวาน)"
            else:
                label = f"{thai_short} ({d_obj.day} {m_txt} {byear})"
            res.append({
                "iso": iso_d,
                "thai": thai_short,
                "label": label,
            })
        return res

    def _build_receipts_where(
        self,
        status_filter: str = "ALL",
        search: str = "",
        date_prefix: str = "",
        date_from: str = "",
        date_to: str = "",
    ) -> Tuple[str, List[Any]]:
        where_sql = " WHERE 1=1"
        params: List[Any] = []

        # Status filter
        if status_filter and status_filter not in ("ALL", "ทั้งหมด"):
            st_clean = status_filter.upper()
            if "VOID" in st_clean or "ยกเลิก" in status_filter:
                where_sql += " AND status IN ('VOIDED', 'CANCELLED')"
            elif "ISSUED" in st_clean or "ปกติ" in status_filter:
                where_sql += " AND status = 'ISSUED'"
            else:
                where_sql += " AND status = ?"
                params.append(status_filter)

        # Date filtering: Range or Preset
        today = datetime.date.today()
        d_start: Optional[datetime.date] = None
        d_end: Optional[datetime.date] = None

        if date_prefix:
            dp_clean = date_prefix.strip().upper()
            if dp_clean in ("TODAY", "วันนี้"):
                d_start = today
                d_end = today
            elif dp_clean in ("YESTERDAY", "เมื่อวาน"):
                d_start = today - datetime.timedelta(days=1)
                d_end = d_start
            elif dp_clean in ("THIS_WEEK", "7 วันล่าสุด", "WEEK"):
                d_start = today - datetime.timedelta(days=6)
                d_end = today
            elif dp_clean in ("THIS_MONTH", "เดือนนี้", "MONTH"):
                d_start = datetime.date(today.year, today.month, 1)
                d_end = today
            elif dp_clean in ("ALL", "ทั้งหมด", "ทุกวัน"):
                pass
            else:
                # Try parsing as exact date
                parsed_d = parse_date_robust(date_prefix)
                if parsed_d:
                    d_start = parsed_d
                    d_end = parsed_d

        # Explicit date_from / date_to override or set bounds
        if date_from and str(date_from).strip():
            parsed_from = parse_date_robust(date_from)
            if parsed_from:
                d_start = parsed_from
        if date_to and str(date_to).strip():
            parsed_to = parse_date_robust(date_to)
            if parsed_to:
                d_end = parsed_to

        if d_start and d_end and d_start == d_end:
            where_sql += " AND sqlite_normalize_date(receipt_date) = ?"
            params.append(d_start.isoformat())
        else:
            if d_start:
                where_sql += " AND sqlite_normalize_date(receipt_date) >= ?"
                params.append(d_start.isoformat())
            if d_end:
                where_sql += " AND sqlite_normalize_date(receipt_date) <= ?"
                params.append(d_end.isoformat())

        # If date_prefix was a partial string (e.g. YYYY-MM) that didn't resolve to a full date
        dp = str(date_prefix or "").strip()
        if not d_start and not d_end and dp and dp not in ("ALL", "ทั้งหมด", "ทุกวัน"):
            where_sql += " AND (sqlite_normalize_date(receipt_date) LIKE ? OR receipt_date LIKE ? OR sqlite_thai_date(receipt_date) LIKE ?)"
            params.extend([f"%{dp}%", f"%{dp}%", f"%{dp}%"])

        # Universal search: matches receipt_no, student_code, student_name, class_group, department, major, and date!
        search_clean = (search or "").strip()
        if search_clean:
            q = f"%{search_clean}%"
            where_sql += (
                " AND (receipt_no LIKE ? OR student_code LIKE ? OR snapshot_student_name LIKE ?"
                " OR snapshot_class_group LIKE ? OR snapshot_department LIKE ? OR snapshot_major LIKE ?"
                " OR receipt_date LIKE ? OR sqlite_thai_date(receipt_date) LIKE ?)"
            )
            params.extend([q, q, q, q, q, q, q, q])

        return where_sql, params

    def get_receipts_stats(
        self,
        status_filter: str = "ALL",
        search: str = "",
        date_prefix: str = "",
        date_from: str = "",
        date_to: str = "",
    ) -> Dict[str, Any]:
        """Fast SQL aggregation for total count, valid ISSUED count, and net total amount across 100k+ receipts."""
        where_sql, params = self._build_receipts_where(status_filter, search, date_prefix, date_from, date_to)
        sql = (
            "SELECT COUNT(*) as total_count, "
            "COALESCE(SUM(CASE WHEN status = 'ISSUED' THEN 1 ELSE 0 END), 0) as valid_count, "
            "COALESCE(SUM(CASE WHEN status IN ('VOIDED', 'CANCELLED') THEN 1 ELSE 0 END), 0) as void_count, "
            "COALESCE(SUM(CASE WHEN status = 'ISSUED' THEN net_total_amount ELSE 0.0 END), 0.0) as total_amount "
            f"FROM receipts{where_sql}"
        )
        with self._connect() as conn:
            row = conn.execute(sql, params).fetchone()
        return {
            "total_count": int(row["total_count"] or 0),
            "valid_count": int(row["valid_count"] or 0),
            "void_count": int(row["void_count"] or 0),
            "total_amount": round(float(row["total_amount"] or 0.0), 2),
        }

    def get_receipts(
        self,
        status_filter: str = "ALL",
        search: str = "",
        date_prefix: str = "",
        date_from: str = "",
        date_to: str = "",
        limit: Optional[int] = None,
        offset: int = 0,
        include_items: bool = True,
    ) -> List[Dict[str, Any]]:
        """
        Fetches receipts with zero N+1 queries by batch-loading receipt_items in bulk.
        Supports optional LIMIT/OFFSET pagination for instant UI rendering.
        """
        where_sql, params = self._build_receipts_where(status_filter, search, date_prefix, date_from, date_to)
        query = f"SELECT * FROM receipts{where_sql} ORDER BY id DESC"
        if limit is not None and limit > 0:
            query += " LIMIT ? OFFSET ?"
            params.extend([int(limit), max(0, int(offset))])

        with self._connect() as conn:
            rows = conn.execute(query, params).fetchall()
            if not rows:
                return []

            items_by_rcp: Dict[int, List[Dict[str, Any]]] = {}
            if include_items:
                rcp_ids = [r["id"] for r in rows]
                if len(rcp_ids) <= 900:
                    placeholders = ",".join("?" for _ in rcp_ids)
                    item_rows = conn.execute(
                        f"SELECT * FROM receipt_items WHERE receipt_id IN ({placeholders}) ORDER BY receipt_id ASC, line_no ASC",
                        rcp_ids,
                    ).fetchall()
                elif len(rcp_ids) <= 5000:
                    item_rows = []
                    for i in range(0, len(rcp_ids), 900):
                        chunk = rcp_ids[i : i + 900]
                        placeholders = ",".join("?" for _ in chunk)
                        item_rows.extend(
                            conn.execute(
                                f"SELECT * FROM receipt_items WHERE receipt_id IN ({placeholders}) ORDER BY receipt_id ASC, line_no ASC",
                                chunk,
                            ).fetchall()
                        )
                else:
                    # Single full scan is much faster than many IN queries when fetching >5,000 receipts
                    rcp_id_set = set(rcp_ids)
                    all_item_rows = conn.execute(
                        "SELECT * FROM receipt_items ORDER BY receipt_id ASC, line_no ASC"
                    ).fetchall()
                    item_rows = [ir for ir in all_item_rows if ir["receipt_id"] in rcp_id_set]

                for it in item_rows:
                    rid = it["receipt_id"]
                    if rid not in items_by_rcp:
                        items_by_rcp[rid] = []
                    items_by_rcp[rid].append(dict(it))

            result = []
            for r in rows:
                d = dict(r)
                items = items_by_rcp.get(d["id"], [])
                d["items"] = items
                d["item_summary"] = ", ".join(f"{it['item_name']} ({it['line_total']:,.0f})" for it in items)
                d["tx_code"] = d["receipt_no"]
                d["tx_date"] = d["receipt_date"]
                d["tx_type"] = "INCOME"
                d["category"] = d["snapshot_department"]
                d["title"] = f"{d['snapshot_student_name']} - {d['item_summary']}"
                d["amount"] = d["net_total_amount"]
                d["ref_code"] = d["student_code"]
                d["recorded_by"] = d["cashier_name"]
                d["full_name"] = d["snapshot_student_name"]
                d["class_group"] = d["snapshot_class_group"]
                d["faculty"] = d["snapshot_department"]
                d["major"] = d["snapshot_major"]
                result.append(d)
            return result

    def get_receipt_by_no(self, receipt_no: str) -> Optional[Dict[str, Any]]:
        clean_no = (receipt_no or "").strip()
        if not clean_no:
            return None
        with self._connect() as conn:
            r = conn.execute("SELECT * FROM receipts WHERE receipt_no = ? LIMIT 1", (clean_no,)).fetchone()
            if not r:
                return None
            d = dict(r)
            items = [
                dict(it)
                for it in conn.execute(
                    "SELECT * FROM receipt_items WHERE receipt_id = ? ORDER BY line_no ASC",
                    (d["id"],),
                ).fetchall()
            ]
            d["items"] = items
            d["item_summary"] = ", ".join(f"{it['item_name']} ({it['line_total']:,.0f})" for it in items)
            d["tx_code"] = d["receipt_no"]
            d["tx_date"] = d["receipt_date"]
            d["tx_type"] = "INCOME"
            d["category"] = d["snapshot_department"]
            d["title"] = f"{d['snapshot_student_name']} - {d['item_summary']}"
            d["amount"] = d["net_total_amount"]
            d["ref_code"] = d["student_code"]
            d["recorded_by"] = d["cashier_name"]
            d["full_name"] = d["snapshot_student_name"]
            d["class_group"] = d["snapshot_class_group"]
            d["faculty"] = d["snapshot_department"]
            d["major"] = d["snapshot_major"]
            return d

    def void_receipt(self, receipt_no: str, reason: str, actor: str, role: str) -> bool:
        clean_no = (receipt_no or "").strip()
        clean_reason = (reason or "").strip()
        if not clean_no:
            raise ValueError("กรุณาระบุเลขที่ใบเสร็จที่ต้องการยกเลิก")
        if not clean_reason:
            raise ValueError("กรุณาระบุเหตุผลในการยกเลิกใบเสร็จ")

        with self._connect() as conn:
            rcp = conn.execute("SELECT * FROM receipts WHERE receipt_no = ?", (clean_no,)).fetchone()
            if not rcp:
                raise ValueError(f"ไม่พบเลขที่ใบเสร็จ {clean_no}")
            if rcp["status"] == "VOIDED":
                raise ValueError(f"ใบเสร็จเลขที่ {clean_no} ถูกยกเลิกไปแล้ว")

            now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            conn.execute(
                """
                UPDATE receipts
                SET status = 'VOIDED', void_reason = ?, voided_at = ?, voided_by = ?
                WHERE receipt_no = ?
                """,
                (clean_reason, now_str, actor, clean_no),
            )

            if rcp["student_id"]:
                stu = conn.execute("SELECT * FROM students WHERE id = ?", (rcp["student_id"],)).fetchone()
                if stu:
                    rev_paid = max(0.0, round(float(stu["paid_amount"]) - float(rcp["net_total_amount"]), 2))
                    total_fee = float(stu["total_fee"])
                    new_st = "PAID" if rev_paid >= total_fee else ("PARTIAL" if rev_paid > 0 else "UNPAID")
                    conn.execute(
                        "UPDATE students SET paid_amount = ?, status = ? WHERE id = ?",
                        (rev_paid, new_st, stu["id"]),
                    )
            conn.commit()

        self.data_version += 1
        self.log_audit(actor, role, "RECEIPT_VOID", receipt_no, f"Voided receipt {receipt_no}. Reason: {reason.strip()}")
        return True

    @staticmethod
    def _parse_receipt_date_obj(raw_date: Any) -> Optional[datetime.date]:
        return parse_date_robust(raw_date)

    def get_dashboard_filter_metadata(self) -> Dict[str, List[str]]:
        """Retrieves dynamic dropdown options for dashboard filtering: terms, levels, departments."""
        with self._connect() as conn:
            terms = self.get_options("TERM") or ["1/2569", "2/2569"]

            # Distinct levels from students and receipts
            db_levels = [
                r[0]
                for r in conn.execute(
                    "SELECT DISTINCT education_level FROM students WHERE education_level IS NOT NULL AND education_level != ''"
                ).fetchall()
            ]
            rcp_levels = [
                r[0]
                for r in conn.execute(
                    "SELECT DISTINCT snapshot_education_level FROM receipts WHERE snapshot_education_level IS NOT NULL AND snapshot_education_level != ''"
                ).fetchall()
            ]
            std_levels = ["ปวช.1", "ปวช.2", "ปวช.3", "ปวส.1", "ปวส.2"]
            all_levels: List[str] = []
            for lvl in std_levels + db_levels + rcp_levels:
                lvl_clean = str(lvl).strip()
                if lvl_clean and lvl_clean not in all_levels:
                    all_levels.append(lvl_clean)

            # Distinct departments/majors from options, students, and receipts
            opt_depts = self.get_options("DEPARTMENT") or []
            stu_majors = [
                r[0]
                for r in conn.execute(
                    "SELECT DISTINCT major FROM students WHERE major IS NOT NULL AND major != ''"
                ).fetchall()
            ]
            rcp_majors = [
                r[0]
                for r in conn.execute(
                    "SELECT DISTINCT snapshot_major FROM receipts WHERE snapshot_major IS NOT NULL AND snapshot_major != ''"
                ).fetchall()
            ]
            all_depts: List[str] = []
            for d in opt_depts + stu_majors + rcp_majors:
                d_clean = str(d).strip()
                if d_clean and d_clean not in all_depts:
                    all_depts.append(d_clean)
            all_depts.sort()

            return {
                "terms": terms,
                "levels": all_levels,
                "departments": all_depts,
            }

    def get_dashboard_analytics(
        self,
        term_filter: str = "ALL",
        period_filter: str = "ALL",
        level_filter: str = "ALL",
        department_filter: str = "ALL",
    ) -> Dict[str, Any]:
        """
        Comprehensive financial & operational analytics for the Executive Dashboard.
        Supports filtering by term ('ALL', '1/2569'), period ('ALL', 'TODAY', 'THIS_WEEK', 'THIS_MONTH'),
        education level ('ALL', 'ปวช.1', 'ปวส.1', etc.), and department ('ALL', 'การบัญชี', etc.).
        """
        today = datetime.date.today()
        week_ago = today - datetime.timedelta(days=7)

        # 1. Base Query for Receipts with Term, Level, and Department Filters
        sql = "SELECT * FROM receipts WHERE 1=1"
        params: List[Any] = []

        if term_filter and term_filter not in ("ALL", "ทั้งหมด", ""):
            sql += " AND term_code = ?"
            params.append(term_filter.strip())

        if level_filter and level_filter not in ("ALL", "ทั้งหมด", "ทุกระดับชั้น", ""):
            lvl_clean = level_filter.strip()
            sql += " AND (snapshot_education_level = ? OR snapshot_education_level LIKE ? OR snapshot_class_group LIKE ?)"
            params.extend([lvl_clean, f"%{lvl_clean}%", f"%{lvl_clean}%"])

        if department_filter and department_filter not in ("ALL", "ทั้งหมด", "ทุกแผนกวิชา", ""):
            dept_clean = department_filter.strip()
            sql += " AND (snapshot_department = ? OR snapshot_major = ? OR snapshot_department LIKE ? OR snapshot_major LIKE ? OR snapshot_class_group LIKE ?)"
            params.extend([dept_clean, dept_clean, f"%{dept_clean}%", f"%{dept_clean}%", f"%{dept_clean}%"])

        sql += " ORDER BY id DESC"

        # 2. Student Aggregations Filtered by Level and Department
        stu_sql = """
            SELECT 
                COUNT(*) as total_cnt,
                COALESCE(SUM(total_fee), 0.0) as tf,
                COALESCE(SUM(paid_amount), 0.0) as pa,
                COALESCE(SUM(CASE WHEN status = 'PAID' THEN 1 ELSE 0 END), 0) as paid_cnt,
                COALESCE(SUM(CASE WHEN status = 'PARTIAL' THEN 1 ELSE 0 END), 0) as partial_cnt,
                COALESCE(SUM(CASE WHEN status = 'UNPAID' THEN 1 ELSE 0 END), 0) as unpaid_cnt
            FROM students
            WHERE 1=1
        """
        stu_params: List[Any] = []

        if level_filter and level_filter not in ("ALL", "ทั้งหมด", "ทุกระดับชั้น", ""):
            lvl_clean = level_filter.strip()
            stu_sql += " AND (education_level = ? OR education_level LIKE ? OR class_group LIKE ?)"
            stu_params.extend([lvl_clean, f"%{lvl_clean}%", f"%{lvl_clean}%"])

        if department_filter and department_filter not in ("ALL", "ทั้งหมด", "ทุกแผนกวิชา", ""):
            dept_clean = department_filter.strip()
            stu_sql += " AND (major = ? OR faculty = ? OR major LIKE ? OR faculty LIKE ? OR class_group LIKE ?)"
            stu_params.extend([dept_clean, dept_clean, f"%{dept_clean}%", f"%{dept_clean}%", f"%{dept_clean}%"])

        with self._connect() as conn:
            all_term_receipts = [dict(r) for r in conn.execute(sql, params).fetchall()]
            stu_row = conn.execute(stu_sql, stu_params).fetchone()

        # 3. Filter Receipts by Period & Calculate Today, Monthly, and All-Time Totals
        period_clean = (period_filter or "ALL").upper()
        filtered_receipts: List[Dict[str, Any]] = []
        today_income = 0.0
        today_receipt_count = 0
        month_income = 0.0
        month_receipt_count = 0
        all_time_income = 0.0
        all_time_receipt_count = 0

        for r in all_term_receipts:
            d = self._parse_receipt_date_obj(r.get("receipt_date"))
            amt = float(r.get("net_total_amount") or 0.0)
            is_valid = (r.get("status") == "ISSUED")

            if is_valid:
                all_time_income += amt
                all_time_receipt_count += 1

                # Track today's stats (for active term/level/dept)
                if d == today:
                    today_income += amt
                    today_receipt_count += 1

                # Track monthly stats (for active term/level/dept)
                if d is not None and d.year == today.year and d.month == today.month:
                    month_income += amt
                    month_receipt_count += 1

            # Match period filter
            if period_clean in ("TODAY", "วันนี้"):
                if d == today:
                    filtered_receipts.append(r)
            elif period_clean in ("THIS_WEEK", "7 วันล่าสุด", "WEEK"):
                if d is not None and d >= week_ago:
                    filtered_receipts.append(r)
            elif period_clean in ("THIS_MONTH", "เดือนนี้", "MONTH"):
                if d is not None and d.year == today.year and d.month == today.month:
                    filtered_receipts.append(r)
            else:
                filtered_receipts.append(r)

        # 4. Aggregations on Filtered Receipts
        total_income = 0.0
        receipt_count = 0
        void_receipt_count = 0
        method_map: Dict[str, Dict[str, Any]] = {}
        level_map: Dict[str, Dict[str, Any]] = {}
        dept_map: Dict[str, Dict[str, Any]] = {}
        valid_receipt_ids: List[int] = []

        for r in filtered_receipts:
            status = r.get("status", "ISSUED")
            amt = float(r.get("net_total_amount") or 0.0)

            if status == "ISSUED":
                total_income += amt
                receipt_count += 1
                valid_receipt_ids.append(r["id"])

                # Group by payment method
                pm = r.get("payment_method") or "เงินสด (Cash)"
                if pm not in method_map:
                    method_map[pm] = {"name": pm, "amount": 0.0, "count": 0}
                method_map[pm]["amount"] += amt
                method_map[pm]["count"] += 1

                # Group by education level
                lvl = r.get("snapshot_education_level") or "ปวช.1"
                if lvl not in level_map:
                    level_map[lvl] = {"level": lvl, "amount": 0.0, "count": 0}
                level_map[lvl]["amount"] += amt
                level_map[lvl]["count"] += 1

                # Group by department/major
                dept = r.get("snapshot_major") or r.get("snapshot_department") or "ทั่วไป"
                if dept not in dept_map:
                    dept_map[dept] = {"name": dept, "amount": 0.0, "count": 0}
                dept_map[dept]["amount"] += amt
                dept_map[dept]["count"] += 1

            elif status == "VOIDED":
                void_receipt_count += 1

        total_income = round(total_income, 2)
        today_income = round(today_income, 2)
        month_income = round(month_income, 2)
        all_time_income = round(all_time_income, 2)
        avg_receipt_amount = round(total_income / receipt_count, 2) if receipt_count > 0 else 0.0

        # Build list for payment methods
        by_payment_method: List[Dict[str, Any]] = []
        for m_data in method_map.values():
            pct = round((m_data["amount"] / total_income * 100.0), 1) if total_income > 0 else 0.0
            by_payment_method.append({
                "name": m_data["name"],
                "amount": round(m_data["amount"], 2),
                "count": m_data["count"],
                "percent": pct,
            })
        by_payment_method.sort(key=lambda x: x["amount"], reverse=True)

        # Build list for education levels
        by_education_level: List[Dict[str, Any]] = []
        for l_data in level_map.values():
            pct = round((l_data["amount"] / total_income * 100.0), 1) if total_income > 0 else 0.0
            by_education_level.append({
                "level": l_data["level"],
                "amount": round(l_data["amount"], 2),
                "count": l_data["count"],
                "percent": pct,
            })
        by_education_level.sort(key=lambda x: x["amount"], reverse=True)

        # Build list for departments
        by_department: List[Dict[str, Any]] = []
        for d_data in dept_map.values():
            pct = round((d_data["amount"] / total_income * 100.0), 1) if total_income > 0 else 0.0
            by_department.append({
                "name": d_data["name"],
                "amount": round(d_data["amount"], 2),
                "count": d_data["count"],
                "percent": pct,
            })
        by_department.sort(key=lambda x: x["amount"], reverse=True)

        # 5. Top 5 Fee Items
        top_fee_items: List[Dict[str, Any]] = []
        if valid_receipt_ids:
            with self._connect() as conn:
                item_counts: Dict[str, Dict[str, Any]] = {}
                for i in range(0, len(valid_receipt_ids), 900):
                    chunk = valid_receipt_ids[i : i + 900]
                    placeholders = ",".join("?" for _ in chunk)
                    fee_sql = f"""
                        SELECT item_name, 
                               COALESCE(SUM(quantity), 0) as qty, 
                               COALESCE(SUM(line_total), 0.0) as total
                        FROM receipt_items
                        WHERE receipt_id IN ({placeholders})
                        GROUP BY item_name
                    """
                    for fr in conn.execute(fee_sql, chunk).fetchall():
                        iname = fr["item_name"]
                        if iname not in item_counts:
                            item_counts[iname] = {"item_name": iname, "qty": 0.0, "total": 0.0}
                        item_counts[iname]["qty"] += float(fr["qty"])
                        item_counts[iname]["total"] += float(fr["total"])

                sorted_fees = sorted(item_counts.values(), key=lambda x: x["total"], reverse=True)[:5]
                for fr in sorted_fees:
                    f_total = round(float(fr["total"]), 2)
                    f_pct = round((f_total / total_income * 100.0), 1) if total_income > 0 else 0.0
                    top_fee_items.append({
                        "item_name": fr["item_name"],
                        "quantity": int(fr["qty"]) if float(fr["qty"]).is_integer() else float(fr["qty"]),
                        "total": f_total,
                        "percent": f_pct,
                    })

        # 6. Student & Tuition Stats
        total_stu = int(stu_row["total_cnt"] or 0)
        tf = round(float(stu_row["tf"] or 0.0), 2)
        pa = round(float(stu_row["pa"] or 0.0), 2)
        paid_stu = int(stu_row["paid_cnt"] or 0)
        partial_stu = int(stu_row["partial_cnt"] or 0)
        unpaid_stu = int(stu_row["unpaid_cnt"] or 0)
        outstanding_ar = max(0.0, round(tf - pa, 2))
        collection_rate = round((pa / tf * 100.0), 1) if tf > 0 else 0.0

        # Recent 15 receipts
        recent_receipts = filtered_receipts[:15]

        period_labels = {
            "ALL": "ข้อมูลทั้งหมด",
            "TODAY": "วันนี้",
            "THIS_WEEK": "7 วันล่าสุด",
            "THIS_MONTH": "เดือนนี้",
        }
        active_period_label = period_labels.get(period_clean, period_clean)

        thai_months = [
            "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
            "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
        ]
        thai_month_name = thai_months[today.month] if 1 <= today.month <= 12 else ""
        thai_year_num = today.year + 543
        month_label_thai = f"{thai_month_name} {thai_year_num}"

        return {
            "term_filter": term_filter,
            "period_filter": period_filter,
            "period_label": active_period_label,
            "level_filter": level_filter,
            "department_filter": department_filter,
            "level_label": level_filter if level_filter not in ("ALL", "ทั้งหมด", "") else "ทุกระดับชั้น",
            "department_label": department_filter if department_filter not in ("ALL", "ทั้งหมด", "") else "ทุกแผนกวิชา",
            "total_income": total_income,
            "receipt_count": receipt_count,
            "void_receipt_count": void_receipt_count,
            "today_income": today_income,
            "today_receipt_count": today_receipt_count,
            "month_income": month_income,
            "month_receipt_count": month_receipt_count,
            "month_label_thai": month_label_thai,
            "all_time_income": all_time_income,
            "all_time_receipt_count": all_time_receipt_count,
            "avg_receipt_amount": avg_receipt_amount,
            "student_count": total_stu,
            "paid_student_count": paid_stu,
            "partial_student_count": partial_stu,
            "unpaid_student_count": unpaid_stu,
            "total_tuition_expected": tf,
            "total_tuition_collected": pa,
            "outstanding_receivables": outstanding_ar,
            "collection_rate_percent": collection_rate,
            "by_payment_method": by_payment_method,
            "by_education_level": by_education_level,
            "by_department": by_department,
            "top_fee_items": top_fee_items,
            "recent_receipts": recent_receipts,
        }

    def get_dashboard_summary(self) -> Dict[str, Any]:
        with self._connect() as conn:
            inc_row = conn.execute(
                "SELECT COUNT(*) as cnt, COALESCE(SUM(net_total_amount), 0.0) as total FROM receipts WHERE status = 'ISSUED'"
            ).fetchone()
            void_cnt = conn.execute(
                "SELECT COUNT(*) as cnt FROM receipts WHERE status = 'VOIDED'"
            ).fetchone()["cnt"]
            stu_row = conn.execute(
                "SELECT COUNT(*) as cnt, COALESCE(SUM(total_fee), 0.0) as tf, COALESCE(SUM(paid_amount), 0.0) as pa FROM students"
            ).fetchone()
            unpaid_cnt = conn.execute(
                "SELECT COUNT(*) as cnt FROM students WHERE status != 'PAID'"
            ).fetchone()["cnt"]

        total_income = float(inc_row["total"])
        total_fee = float(stu_row["tf"])
        paid_fee = float(stu_row["pa"])
        outstanding_ar = max(0.0, round(total_fee - paid_fee, 2))

        return {
            "total_income": total_income,
            "total_expense": 0.0,
            "net_balance": total_income,
            "receipt_count": inc_row["cnt"],
            "void_receipt_count": void_cnt,
            "student_count": stu_row["cnt"],
            "unpaid_student_count": unpaid_cnt,
            "total_tuition_expected": total_fee,
            "total_tuition_collected": paid_fee,
            "outstanding_receivables": outstanding_ar,
            "budget_allocated": total_fee,
            "budget_spent": paid_fee,
            "pending_requisitions": unpaid_cnt,
        }

    def get_daily_closing_summary(
        self,
        date_prefix: str = None,
        date_from: str = "",
        date_to: str = "",
    ) -> Dict[str, Any]:
        target_date = date_prefix if date_prefix is not None else datetime.datetime.now().strftime("%Y-%m-%d")
        receipts = self.get_receipts(
            status_filter="ISSUED",
            date_prefix=target_date,
            date_from=date_from,
            date_to=date_to,
        )

        income_total = 0.0
        by_method: Dict[str, float] = {}
        by_fee_item: Dict[str, float] = {}

        for r in receipts:
            amt = float(r.get("net_total_amount") or 0.0)
            method = r.get("payment_method") or "เงินสด (Cash)"
            income_total += amt
            by_method[method] = round(by_method.get(method, 0.0) + amt, 2)
            for it in r.get("items", []):
                iname = it.get("item_name") or "ค่าธรรมเนียมการศึกษา"
                by_fee_item[iname] = round(by_fee_item.get(iname, 0.0) + float(it.get("line_total") or 0.0), 2)

        display_date = target_date or "ทั้งหมด"
        if date_from and date_to:
            display_date = f"{format_receipt_date_thai(date_from)} - {format_receipt_date_thai(date_to)}"
        elif target_date and target_date not in ("ALL", "ทั้งหมด"):
            display_date = format_receipt_date_thai(target_date)

        return {
            "date": display_date,
            "tx_count": len(receipts),
            "income_total": round(income_total, 2),
            "expense_total": 0.0,
            "net_daily": round(income_total, 2),
            "income_by_method": by_method,
            "income_by_fee_item": by_fee_item,
            "transactions": receipts,
        }


    # ------------------------------------------------------------------
    # Student CRUD (Create, Read, Update, Delete, + Bulk Excel/CSV Import)
    # ------------------------------------------------------------------
    def get_distinct_student_years(self) -> List[str]:
        """Returns available admission / cohort years, e.g. ['ปี 2569 (รุ่น 69)', 'ปี 2568 (รุ่น 68)', ...]"""
        found = set()
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT DISTINCT substr(student_code, 1, 2) AS yr FROM students WHERE length(student_code) >= 2 ORDER BY yr DESC"
            ).fetchall()
            for r in rows:
                pfx = str(r["yr"]).strip()
                if pfx.isdigit() and len(pfx) == 2:
                    found.add(pfx)
        for d in ["69", "68", "67", "66"]:
            found.add(d)
        sorted_pfx = sorted(list(found), reverse=True)
        return [f"ปี 25{p} (รุ่น {p})" for p in sorted_pfx]

    def _build_students_where(
        self,
        search: str = "",
        status_filter: str = "ALL",
        level_filter: str = "ALL",
        dept_filter: str = "ALL",
        major_filter: str = "ALL",
        year_level_filter: str = "ALL",
        admission_year_filter: str = "ALL",
        year_filter: str = "ALL",
    ) -> Tuple[str, List[Any]]:
        where_sql = " WHERE 1=1"
        params: List[Any] = []
        search_clean = (search or "").strip()
        if search_clean:
            q = f"%{search_clean}%"
            where_sql += " AND (student_code LIKE ? OR full_name LIKE ? OR faculty LIKE ? OR major LIKE ? OR class_group LIKE ?)"
            params.extend([q, q, q, q, q])
        if status_filter and status_filter not in ("ALL", "ทั้งหมด"):
            where_sql += " AND status = ?"
            params.append(status_filter)
        if level_filter and level_filter not in ("ALL", "ทุกระดับชั้น", "ทุกชั้นปี"):
            where_sql += " AND (class_group LIKE ? OR education_level LIKE ?)"
            params.extend([f"%{level_filter}%", f"%{level_filter}%"])
        if dept_filter and dept_filter not in ("ALL", "ทุกคณะ", "ทุกแผนกวิชา"):
            where_sql += " AND (faculty = ? OR major = ?)"
            params.extend([dept_filter, dept_filter])
        if major_filter and major_filter not in ("ALL", "ทุกสาขาวิชา", "ทุกแผนกวิชา"):
            where_sql += " AND major = ?"
            params.append(major_filter)

        # Filter by year level (ชั้นปี เช่น ปี 1, ปี 2, ปี 3, ปี 4)
        eff_yr_lvl = year_level_filter
        if (not eff_yr_lvl or eff_yr_lvl in ("ALL", "ทั้งหมด")) and year_filter and year_filter not in ("ALL", "ทั้งหมด"):
            eff_yr_lvl = year_filter

        if eff_yr_lvl and eff_yr_lvl not in ("ALL", "ทั้งหมด", "ทุกชั้นปี"):
            m = re.search(r"(\d+)", str(eff_yr_lvl))
            if m:
                target_yr = int(m.group(1))
                where_sql += " AND (year_level = ? OR (COALESCE(year_level, 0) = 0 AND (class_group LIKE ? OR education_level LIKE ?)))"
                params.extend([target_yr, f"%{target_yr}/%", f"%.{target_yr}%"])
            else:
                where_sql += " AND (class_group LIKE ? OR education_level LIKE ?)"
                params.extend([f"%{eff_yr_lvl}%", f"%{eff_yr_lvl}%"])

        # Filter by admission / cohort year (ปีที่เข้าศึกษา / รุ่น เช่น ปี 2569 (รุ่น 69) -> รหัสขึ้นต้น 69)
        if admission_year_filter and admission_year_filter not in ("ALL", "ทั้งหมด", "ทุกรุ่น", "ทุกปีการศึกษา"):
            digits = re.findall(r"\d+", str(admission_year_filter))
            cohort = None
            for d in digits:
                if len(d) == 4 and (d.startswith("25") or d.startswith("20")):
                    cohort = d[-2:]
                    break
                elif len(d) == 2:
                    cohort = d
                    break
            if not cohort and digits:
                cohort = digits[0]
            if cohort:
                where_sql += " AND student_code LIKE ?"
                params.append(f"{cohort}%")

        return where_sql, params

    def count_students(
        self,
        search: str = "",
        status_filter: str = "ALL",
        level_filter: str = "ALL",
        dept_filter: str = "ALL",
        major_filter: str = "ALL",
        year_level_filter: str = "ALL",
        admission_year_filter: str = "ALL",
        year_filter: str = "ALL",
    ) -> int:
        where_sql, params = self._build_students_where(
            search=search,
            status_filter=status_filter,
            level_filter=level_filter,
            dept_filter=dept_filter,
            major_filter=major_filter,
            year_level_filter=year_level_filter,
            admission_year_filter=admission_year_filter,
            year_filter=year_filter,
        )
        with self._connect() as conn:
            row = conn.execute(f"SELECT COUNT(*) as c FROM students{where_sql}", params).fetchone()
            return int(row["c"] or 0)

    def get_student_by_code(self, student_code: str) -> Optional[Dict[str, Any]]:
        clean = (student_code or "").strip()
        if not clean:
            return None
        with self._connect() as conn:
            r = conn.execute("SELECT * FROM students WHERE student_code = ? LIMIT 1", (clean,)).fetchone()
            if not r:
                return None
            d = dict(r)
            d.setdefault("education_level", "ปวช.1")
            d.setdefault("class_group", "ปวช.1/1")
            d.setdefault("registered_credits", 20)
            d["balance"] = max(0.0, round(float(d["total_fee"]) - float(d["paid_amount"]), 2))
            return d

    def get_students_by_codes(self, student_codes: List[str]) -> List[Dict[str, Any]]:
        clean_codes = [str(c).strip() for c in student_codes if str(c).strip()]
        if not clean_codes:
            return []
        result = []
        with self._connect() as conn:
            for i in range(0, len(clean_codes), 900):
                chunk = clean_codes[i : i + 900]
                placeholders = ",".join("?" for _ in chunk)
                rows = conn.execute(
                    f"SELECT * FROM students WHERE student_code IN ({placeholders}) ORDER BY student_code ASC",
                    chunk,
                ).fetchall()
                for r in rows:
                    d = dict(r)
                    d.setdefault("education_level", "ปวช.1")
                    d.setdefault("class_group", "ปวช.1/1")
                    d.setdefault("registered_credits", 20)
                    d["balance"] = max(0.0, round(float(d["total_fee"]) - float(d["paid_amount"]), 2))
                    result.append(d)
        return result

    def get_students(
        self,
        search: str = "",
        status_filter: str = "ALL",
        level_filter: str = "ALL",
        dept_filter: str = "ALL",
        major_filter: str = "ALL",
        limit: Optional[int] = None,
        offset: int = 0,
        year_level_filter: str = "ALL",
        admission_year_filter: str = "ALL",
        year_filter: str = "ALL",
    ) -> List[Dict[str, Any]]:
        where_sql, params = self._build_students_where(
            search=search,
            status_filter=status_filter,
            level_filter=level_filter,
            dept_filter=dept_filter,
            major_filter=major_filter,
            year_level_filter=year_level_filter,
            admission_year_filter=admission_year_filter,
            year_filter=year_filter,
        )
        query = f"SELECT * FROM students{where_sql} ORDER BY student_code ASC"
        if limit is not None and limit > 0:
            query += " LIMIT ? OFFSET ?"
            params.extend([int(limit), max(0, int(offset))])

        with self._connect() as conn:
            rows = conn.execute(query, params).fetchall()
            result = []
            for r in rows:
                d = dict(r)
                d.setdefault("education_level", "ปวช.1")
                d.setdefault("class_group", "ปวช.1/1")
                d.setdefault("registered_credits", 20)
                d["balance"] = max(0.0, round(float(d["total_fee"]) - float(d["paid_amount"]), 2))
                result.append(d)
            return result

    def add_student(
        self,
        student_code: str,
        full_name: str,
        citizen_id: str = "",
        faculty: str = "บริหารธุรกิจ / พาณิชยกรรม",
        major: str = "การบัญชี",
        year_level: int = 1,
        scholarship_type: str = "ปกติ (Regular)",
        total_fee: float = 950.0,
        phone: str = "",
        actor: str = "admin",
        role: str = "ADMIN",
        education_level: str = "ปวช.1",
        class_group: str = "ปวช.1/1",
        registered_credits: int = 20,
    ) -> bool:
        code = (student_code or "").strip()
        name = (full_name or "").strip()
        if not code or not name:
            raise ValueError("กรุณากรอกรหัสประจำตัวนักศึกษาและชื่อ-นามสกุล")
        status = "PAID" if total_fee <= 0 else "UNPAID"
        derived_yr, derived_edu = derive_year_level_and_education_level(
            class_group, default_year=int(year_level), default_edu="ปวช.1"
        )
        eff_yr = int(year_level) if year_level != 1 else derived_yr
        eff_edu = derived_edu if (education_level in ("", "ปวช.1") and derived_edu) else (education_level.strip() or derived_edu)
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO students
                (student_code, full_name, citizen_id, education_level, class_group, faculty, major, year_level, scholarship_type, registered_credits, total_fee, paid_amount, status, phone)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0.0, ?, ?)
                """,
                (
                    code,
                    name,
                    (citizen_id or "").strip(),
                    eff_edu,
                    (class_group or "ปวช.1/1").strip(),
                    (faculty or "บริหารธุรกิจ / พาณิชยกรรม").strip(),
                    (major or "การบัญชี").strip(),
                    int(eff_yr),
                    scholarship_type,
                    int(registered_credits),
                    float(total_fee),
                    status,
                    (phone or "").strip(),
                ),
            )
            conn.execute("INSERT OR IGNORE INTO master_options (category, option_value) VALUES ('FACULTY', ?)", (faculty.strip(),))
            conn.execute("INSERT OR IGNORE INTO master_options (category, option_value) VALUES ('DEPARTMENT', ?)", (major.strip(),))
            conn.execute("INSERT OR IGNORE INTO master_options (category, option_value) VALUES ('CLASS_LEVEL', ?)", (class_group.strip(),))
            conn.commit()
        for cat in ("FACULTY", "DEPARTMENT", "CLASS_LEVEL"):
            self._options_cache.pop(cat, None)
        self.data_version += 1
        self.log_audit(actor, role, "STUDENT_CREATE", code, f"Added student {name} [{class_group}] Year {eff_yr}")
        return True

    def update_student(
        self,
        student_code: str,
        full_name: str,
        faculty: str,
        major: str,
        class_group: str,
        year_level: Optional[int] = None,
        education_level: str = "",
        actor: str = "admin",
        role: str = "ADMIN",
    ) -> bool:
        code = (student_code or "").strip()
        if not code:
            raise ValueError("กรุณาเลือกหรือระบุรหัสประจำตัวนักศึกษาที่ต้องการแก้ไข")
        derived_yr, derived_edu = derive_year_level_and_education_level(class_group)
        eff_yr = int(year_level) if year_level is not None else derived_yr
        eff_edu = education_level.strip() if education_level else derived_edu
        with self._connect() as conn:
            cur = conn.execute(
                """
                UPDATE students
                SET full_name = ?, faculty = ?, major = ?, class_group = ?, education_level = ?, year_level = ?
                WHERE student_code = ?
                """,
                (full_name.strip(), faculty.strip(), major.strip(), class_group.strip(), eff_edu, eff_yr, code),
            )
            if cur.rowcount == 0:
                raise ValueError(f"ไม่พบรหัสนักศึกษา {code}")
            conn.commit()
        self.data_version += 1
        self.log_audit(actor, role, "STUDENT_UPDATE", code, f"Updated student {full_name} [{class_group}] Year {eff_yr}")
        return True

    def delete_student(self, student_code: str, actor: str = "admin", role: str = "ADMIN") -> bool:
        code = (student_code or "").strip()
        if not code:
            raise ValueError("กรุณาเลือกนักศึกษาที่ต้องการลบ")
        with self._connect() as conn:
            conn.execute("DELETE FROM students WHERE student_code = ?", (code,))
            conn.commit()
        self.data_version += 1
        self.log_audit(actor, role, "STUDENT_DELETE", code, f"Deleted student {code}")
        return True

    def bulk_import_students(
        self,
        file_path: str,
        actor: str,
        role: str,
        default_faculty: str = "บริหารธุรกิจ / พาณิชยกรรม",
        default_major: str = "การบัญชี",
        default_class_group: str = "ปวช.1/1",
    ) -> Dict[str, Any]:
        """
        Imports students from .csv or .xlsx and returns both counts AND the imported student dicts
        so the caller can immediately issue receipts for the whole imported file if desired!
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"ไม่พบไฟล์: {file_path}")

        ext = os.path.splitext(file_path)[1].lower()
        raw_rows: List[List[Any]] = []

        if ext == ".csv":
            content = None
            for enc in ("utf-8-sig", "utf-8", "cp874", "tis-620"):
                try:
                    with open(file_path, "r", encoding=enc) as f:
                        content = f.read()
                    break
                except (UnicodeDecodeError, UnicodeError):
                    continue
            if content is None:
                with open(file_path, "r", encoding="utf-8-sig", errors="replace") as f:
                    content = f.read()
            import io
            reader = csv.reader(io.StringIO(content))
            for idx, row in enumerate(reader):
                if idx == 0 and any("student" in str(c).lower() or "รหัส" in str(c) or "ชื่อ" in str(c) for c in row):
                    continue
                if row and str(row[0]).strip():
                    raw_rows.append(row)
        elif ext == ".xlsx":
            if openpyxl is None:
                raise RuntimeError("ต้องติดตั้งไลบรารี openpyxl เพื่อนำเข้าไฟล์ .xlsx")
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            ws = wb.active
            for idx, row in enumerate(ws.iter_rows(values_only=True)):
                if not row or row[0] is None:
                    continue
                if idx == 0 and ("student" in str(row[0]).lower() or "รหัส" in str(row[0]) or "ชื่อ" in str(row[0])):
                    continue
                raw_rows.append(list(row))
            wb.close()
        else:
            raise ValueError("รองรับเฉพาะไฟล์นามสกุล .csv หรือ .xlsx เท่านั้น")

        inserted = 0
        updated = 0
        imported_codes: List[str] = []
        seen_faculties = set()
        seen_majors = set()
        seen_classes = set()

        with self._connect() as conn:
            existing_map = {
                r["student_code"]: float(r["paid_amount"])
                for r in conn.execute("SELECT student_code, paid_amount FROM students").fetchall()
            }
            to_insert: List[Tuple[Any, ...]] = []
            to_update: List[Tuple[Any, ...]] = []

            conn.execute("BEGIN TRANSACTION;")
            try:
                for r in raw_rows:
                    code = str(r[0]).strip()
                    name = str(r[1]).strip() if len(r) > 1 and r[1] else "ไม่ระบุชื่อ"
                    cid = str(r[2]).strip() if len(r) > 2 and r[2] else ""
                    fac = str(r[3]).strip() if len(r) > 3 and r[3] else default_faculty
                    maj = str(r[4]).strip() if len(r) > 4 and r[4] else default_major
                    yr_raw = str(r[5]).strip() if len(r) > 5 and r[5] is not None else "1"
                    yr = int(yr_raw) if yr_raw.isdigit() else 1
                    sch = str(r[6]).strip() if len(r) > 6 and r[6] else "ปกติ (Regular)"
                    try:
                        fee = float(r[7]) if len(r) > 7 and r[7] is not None else 950.0
                    except ValueError:
                        fee = 950.0
                    phone = str(r[8]).strip() if len(r) > 8 and r[8] else ""
                    edu_lvl = str(r[9]).strip() if len(r) > 9 and r[9] else default_class_group
                    cls_grp = str(r[10]).strip() if len(r) > 10 and r[10] else default_class_group

                    if code in existing_map:
                        paid = existing_map[code]
                        st_status = "PAID" if paid >= fee else ("PARTIAL" if paid > 0 else "UNPAID")
                        to_update.append((name, cid, edu_lvl, cls_grp, fac, maj, yr, sch, fee, st_status, phone, code))
                        updated += 1
                    else:
                        st_status = "PAID" if fee <= 0 else "UNPAID"
                        to_insert.append((code, name, cid, edu_lvl, cls_grp, fac, maj, yr, sch, fee, st_status, phone))
                        existing_map[code] = 0.0
                        inserted += 1

                    imported_codes.append(code)
                    if fac:
                        seen_faculties.add(fac)
                    if maj:
                        seen_majors.add(maj)
                    if cls_grp:
                        seen_classes.add(cls_grp)

                if to_insert:
                    conn.executemany(
                        """
                        INSERT INTO students
                        (student_code, full_name, citizen_id, education_level, class_group, faculty, major, year_level, scholarship_type, total_fee, paid_amount, status, phone)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0.0, ?, ?)
                        """,
                        to_insert,
                    )
                if to_update:
                    conn.executemany(
                        """
                        UPDATE students
                        SET full_name = ?, citizen_id = ?, education_level = ?, class_group = ?, faculty = ?, major = ?, year_level = ?, scholarship_type = ?, total_fee = ?, status = ?, phone = ?
                        WHERE student_code = ?
                        """,
                        to_update,
                    )

                if seen_faculties:
                    conn.executemany(
                        "INSERT OR IGNORE INTO master_options (category, option_value) VALUES ('FACULTY', ?)",
                        [(f,) for f in seen_faculties],
                    )
                if seen_majors:
                    conn.executemany(
                        "INSERT OR IGNORE INTO master_options (category, option_value) VALUES ('DEPARTMENT', ?)",
                        [(m,) for m in seen_majors],
                    )
                if seen_classes:
                    conn.executemany(
                        "INSERT OR IGNORE INTO master_options (category, option_value) VALUES ('CLASS_LEVEL', ?)",
                        [(c,) for c in seen_classes],
                    )
                conn.commit()
            except Exception:
                conn.rollback()
                raise

        for cat in ("FACULTY", "DEPARTMENT", "CLASS_LEVEL"):
            self._options_cache.pop(cat, None)
        self.data_version += 1
        self.log_audit(
            actor,
            role,
            "STUDENT_BULK_IMPORT",
            os.path.basename(file_path),
            f"Bulk imported {inserted + updated} students ({inserted} new, {updated} updated).",
        )
        return {
            "inserted": inserted,
            "updated": updated,
            "total": inserted + updated,
            "student_codes": imported_codes,
        }

    def get_distinct_student_departments(self) -> List[str]:
        return self.get_options("DEPARTMENT")

    def get_transactions(self, tx_type: str = "ALL", search: str = "") -> List[Dict[str, Any]]:
        return self.get_receipts(search=search)

    def void_transaction(self, tx_code: str, reason: str, actor: str, role: str) -> bool:
        return self.void_receipt(tx_code, reason, actor, role)

    def get_budgets(self) -> List[Dict[str, Any]]:
        return []

    # ------------------------------------------------------------------
    # Course Registration (บัตรลงทะเบียนรายวิชา อศท.28)
    # ------------------------------------------------------------------
    def get_course_catalog(self) -> List[Dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute("SELECT * FROM course_catalog ORDER BY sort_order ASC, id ASC").fetchall()
            return [dict(r) for r in rows]

    def upsert_course_catalog(
        self,
        course_code: str,
        course_name: str,
        theory_hours: int = 0,
        practice_hours: int = 2,
        credits: int = 1,
        major: str = "เทคโนโลยีธุรกิจดิจิทัล",
        class_level: str = "ปวช.1",
        actor: str = "admin",
        role: str = "ADMIN",
    ) -> bool:
        code = (course_code or "").strip()
        name = (course_name or "").strip()
        if not code or not name:
            raise ValueError("กรุณากรอกรหัสวิชาและชื่อวิชา")
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO course_catalog (course_code, course_name, theory_hours, practice_hours, credits, major, class_level)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(course_code) DO UPDATE SET
                    course_name = excluded.course_name,
                    theory_hours = excluded.theory_hours,
                    practice_hours = excluded.practice_hours,
                    credits = excluded.credits,
                    major = excluded.major,
                    class_level = excluded.class_level
                """,
                (code, name, int(theory_hours), int(practice_hours), int(credits), major.strip(), class_level.strip()),
            )
            conn.commit()
        self.data_version += 1
        self.log_audit(actor, role, "COURSE_UPSERT", code, f"Saved course {code} {name} ({theory_hours}-{practice_hours}-{credits})")
        return True

    def get_student_registered_courses(self, student_code: str, term_code: str = "2/2569") -> List[Dict[str, Any]]:
        clean_code = (student_code or "").strip()
        if not clean_code:
            return []
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM student_course_registrations WHERE student_code = ? ORDER BY line_no ASC, id ASC",
                (clean_code,),
            ).fetchall()
            if rows:
                return [dict(r) for r in rows]
            # Default fallback to the 11 SPVC courses from EXofBill so any selected student has a ready template
            cat = conn.execute("SELECT * FROM course_catalog ORDER BY sort_order ASC, id ASC").fetchall()
            return [dict(c) for c in cat]

    def save_student_registered_courses(
        self,
        student_code: str,
        term_code: str,
        courses_list: List[Dict[str, Any]],
        actor: str = "admin",
        role: str = "ADMIN",
    ) -> bool:
        code = (student_code or "").strip()
        if not code:
            raise ValueError("กรุณาเลือกนักศึกษา")
        with self._connect() as conn:
            conn.execute("DELETE FROM student_course_registrations WHERE student_code = ?", (code,))
            for idx, c in enumerate(courses_list, start=1):
                conn.execute(
                    """
                    INSERT INTO student_course_registrations
                    (student_code, term_code, line_no, course_code, course_name, theory_hours, practice_hours, credits)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        code,
                        (term_code or "2/2569").strip(),
                        idx,
                        str(c.get("course_code", "")).strip(),
                        str(c.get("course_name", "")).strip(),
                        int(float(c.get("theory_hours", 0))),
                        int(float(c.get("practice_hours", 0))),
                        int(float(c.get("credits", 0))),
                    ),
                )
            conn.commit()
        self.data_version += 1
        self.log_audit(actor, role, "COURSE_REGISTER", code, f"Registered {len(courses_list)} courses for {code}")
        return True
