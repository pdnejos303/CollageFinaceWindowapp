import datetime
import json
import os
import shutil
import sqlite3
import sys
import zipfile
from typing import Any, Dict

DEFAULT_CONFIG: Dict[str, Any] = {
    "college_name": "วิทยาลัยอาชีวศึกษาสุพรรณบุรี",
    "college_address": "455 ถ.พระพันวษา ต.ท่าพี่เลี้ยง อ.เมืองสุพรรณบุรี จ.สุพรรณบุรี 72000",
    "college_tax_id": "0994000159721",
    "fiscal_year": "2569",
    "current_semester": "1/2569",
    "cashier_position": "ครู",
    "receipt_book_no": "7006",
    "receipt_prefix": "",
    "receipt_no_start": "700605433",
    "receipt_no_end": "700611000",
    "db_engine": "sqlite",
    "database_url": "",
    "webhook_url": "",
    "webhook_enabled": False,
    "webhook_alert_threshold": 50000.0,
    "auto_backup_on_exit": True,
    "auto_lock_minutes": 15,
    "default_export_folder": "",
    "evidence_folder": "",
    "theme_accent": "#000000",
}


def thai_baht_text(amount: float) -> str:
    """Converts a numeric amount into standard Thai Baht text (e.g., 2500.00 -> 'สองพันห้าร้อยบาทถ้วน')."""
    try:
        val = round(float(amount), 2)
    except (ValueError, TypeError):
        return "ศูนย์บาทถ้วน"

    if val <= 0:
        return "ศูนย์บาทถ้วน"

    thai_num = ["ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"]
    thai_pos = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน", "ล้าน"]

    def _convert_int(n: int) -> str:
        if n == 0:
            return ""
        if n >= 1_000_000:
            millions = n // 1_000_000
            rem = n % 1_000_000
            return _convert_int(millions) + "ล้าน" + _convert_int(rem)
        s = str(n)
        length = len(s)
        res = []
        for i, ch in enumerate(s):
            digit = int(ch)
            pos = length - i - 1
            if digit == 0:
                continue
            if pos == 1 and digit == 1:
                res.append("สิบ")
            elif pos == 1 and digit == 2:
                res.append("ยี่สิบ")
            elif pos == 0 and digit == 1 and length > 1:
                res.append("เอ็ด")
            else:
                res.append(thai_num[digit] + thai_pos[pos])
        return "".join(res)

    baht_part = int(val)
    satang_part = int(round((val - baht_part) * 100))

    baht_str = _convert_int(baht_part) if baht_part > 0 else "ศูนย์"
    if satang_part == 0:
        return f"{baht_str}บาทถ้วน"
    return f"{baht_str}บาท{_convert_int(satang_part)}สตางค์"


def format_thai_date_short(dt: datetime.date = None) -> str:
    """Returns Thai Buddhist date in D/M/YYYY format matching EX (e.g., '6/11/2569')."""
    d = dt or datetime.date.today()
    byear = d.year + 543 if d.year < 2400 else d.year
    return f"{d.day}/{d.month}/{byear}"


def get_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.abspath(os.path.dirname(os.path.dirname(__file__)))


def get_config_path(base_dir: str = None) -> str:
    base = base_dir or get_base_dir()
    return os.path.join(base, "config.json")


def load_config(base_dir: str = None) -> Dict[str, Any]:
    base = base_dir or get_base_dir()
    cfg = DEFAULT_CONFIG.copy()
    cfg["default_export_folder"] = os.path.join(base, "exports_output")
    cfg["evidence_folder"] = os.path.join(base, "data", "evidence")
    os.makedirs(cfg["default_export_folder"], exist_ok=True)
    os.makedirs(cfg["evidence_folder"], exist_ok=True)

    path = get_config_path(base)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
                for k, v in loaded.items():
                    if k in cfg and v != "":
                        cfg[k] = v
                    elif k in cfg and isinstance(v, (bool, int, float)):
                        cfg[k] = v
        except Exception as e:
            print(f"Warning: could not load config.json: {e}")
    return cfg


def save_config(cfg: Dict[str, Any], base_dir: str = None) -> bool:
    path = get_config_path(base_dir)
    data = {k: cfg.get(k, DEFAULT_CONFIG.get(k)) for k in DEFAULT_CONFIG.keys()}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    return True


def backup_database(db_path: str, backup_root: str = None) -> str:
    if not os.path.exists(db_path):
        return ""
    base = backup_root or os.path.join(os.path.dirname(db_path), "backups")
    os.makedirs(base, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(base, f"spvc_receipt_backup_{stamp}.db")
    try:
        src_conn = sqlite3.connect(db_path)
        dst_conn = sqlite3.connect(dest)
        with dst_conn:
            src_conn.backup(dst_conn)
        dst_conn.close()
        src_conn.close()
    except Exception:
        shutil.copy2(db_path, dest)
    return dest


def export_full_zip_backup(db_path: str, evidence_dir: str, output_dir: str = None) -> str:
    hot_db = backup_database(db_path)
    out_folder = output_dir or os.path.join(os.path.dirname(db_path), "backups")
    os.makedirs(out_folder, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_path = os.path.join(out_folder, f"SPVC_FullBackup_{stamp}.zip")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        if hot_db and os.path.exists(hot_db):
            zf.write(hot_db, arcname="college_finance.db")
        if evidence_dir and os.path.exists(evidence_dir):
            for root_dir, _, files in os.walk(evidence_dir):
                for file in files:
                    full_p = os.path.join(root_dir, file)
                    rel_p = os.path.relpath(full_p, os.path.dirname(evidence_dir))
                    zf.write(full_p, arcname=rel_p)
    return zip_path


def restore_database(source_backup_path: str, target_db_path: str) -> str:
    if not os.path.exists(source_backup_path):
        raise FileNotFoundError(f"ไม่พบไฟล์สำรองข้อมูล: {source_backup_path}")
    with open(source_backup_path, "rb") as f:
        header = f.read(16)
    if not header.startswith(b"SQLite format 3"):
        raise ValueError("ไฟล์ที่เลือกไม่ใช่ฐานข้อมูล SQLite ที่ถูกต้อง")

    chk_conn = sqlite3.connect(source_backup_path)
    res = chk_conn.execute("PRAGMA integrity_check;").fetchone()
    chk_conn.close()
    if not res or str(res[0]).lower() != "ok":
        raise ValueError("ไฟล์สำรองข้อมูลมีความเสียหาย (Integrity Check Failed)")

    if os.path.exists(target_db_path):
        backup_database(target_db_path)

    for suffix in ("-wal", "-shm"):
        aux = target_db_path + suffix
        if os.path.exists(aux):
            try:
                os.remove(aux)
            except Exception:
                pass

    shutil.copy2(source_backup_path, target_db_path)
    return target_db_path
