import hashlib
import hmac
import math
import os
import re
from typing import Dict, List


ROLE_PERMISSIONS: Dict[str, List[str]] = {
    "ADMIN": [
        "dashboard.view",
        "students.view", "students.manage", "students.pay",
        "ledger.view", "ledger.create", "ledger.void",
        "budget.view", "budget.manage", "budget.approve",
        "payroll.view", "payroll.manage", "payroll.pay",
        "audit.view",
        "settings.manage"
    ],
    "FINANCE": [
        "dashboard.view",
        "students.view", "students.manage", "students.pay",
        "ledger.view", "ledger.create",
        "budget.view", "budget.manage",
        "payroll.view", "payroll.manage", "payroll.pay",
        "audit.view"
    ],
    "AUDITOR": [
        "dashboard.view",
        "students.view",
        "ledger.view",
        "budget.view",
        "payroll.view",
        "audit.view"
    ]
}


class SecurityManager:
    """Handles PBKDF2 password hashing, RBAC permission verification, and PDPA data masking."""

    ITERATIONS = 120_000

    @classmethod
    def hash_password(cls, plain_password: str, salt_hex: str = None) -> str:
        """Returns 'salt_hex$hash_hex' using PBKDF2-HMAC-SHA256."""
        if plain_password is None:
            plain_password = ""
        if not salt_hex:
            salt_bytes = os.urandom(16)
            salt_hex = salt_bytes.hex()
        else:
            salt_bytes = bytes.fromhex(salt_hex)

        dk = hashlib.pbkdf2_hmac(
            "sha256",
            str(plain_password).encode("utf-8"),
            salt_bytes,
            cls.ITERATIONS
        )
        return f"{salt_hex}${dk.hex()}"

    @classmethod
    def verify_password(cls, plain_password: str, stored_hash: str) -> bool:
        """Verifies a plain password against a 'salt_hex$hash_hex' string with constant-time comparison."""
        if not plain_password or not stored_hash or "$" not in stored_hash:
            return False
        salt_hex, _ = stored_hash.split("$", 1)
        candidate = cls.hash_password(plain_password, salt_hex=salt_hex)
        return hmac.compare_digest(candidate, stored_hash)

    @classmethod
    def has_permission(cls, role: str, permission: str) -> bool:
        role_upper = (role or "").strip().upper()
        perms = ROLE_PERMISSIONS.get(role_upper, [])
        return (permission or "").strip() in perms

    @staticmethod
    def mask_citizen_id(citizen_id: str) -> str:
        """Masks a 13-digit Thai Citizen ID for PDPA compliance (e.g. 1-1005-XXXXX-12-3)."""
        if not citizen_id:
            return ""
        digits = re.sub(r"\D", "", str(citizen_id))
        if not digits:
            return ""
        if len(digits) == 13:
            return f"{digits[0]}-{digits[1:5]}-XXXXX-{digits[10:12]}-{digits[12]}"
        if len(digits) >= 6:
            return digits[:3] + "X" * (len(digits) - 5) + digits[-2:]
        return "XXX-XXXX-XXXX"

    @staticmethod
    def mask_bank_account(acc_no: str) -> str:
        """Masks a bank account number for PDPA compliance (e.g. XXX-X-XX456-7)."""
        if not acc_no:
            return ""
        digits = re.sub(r"\D", "", str(acc_no))
        if not digits:
            return ""
        if len(digits) >= 10:
            return f"XXX-X-XX{digits[-4:-1]}-{digits[-1]}"
        if len(digits) >= 4:
            return "X" * (len(digits) - 4) + digits[-4:]
        return "XXX-X-XXXXX-X"

    @staticmethod
    def format_currency(amount: float) -> str:
        """Formats a float as Thai Baht string (e.g. ฿125,400.00)."""
        try:
            val = float(amount)
            if math.isnan(val) or math.isinf(val):
                return "฿0.00"
            return f"฿{val:,.2f}"
        except (ValueError, TypeError):
            return "฿0.00"
