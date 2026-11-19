from abc import ABC, abstractmethod
from typing import Any, Dict


def safe_float(val: Any, default: float = 0.0) -> float:
    """Safely convert any value to float, handling None, empty string, or invalid formats."""
    if val is None or val == "":
        return default
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


def safe_int(val: Any, default: int = 0) -> int:
    """Safely convert any value to int, handling None, empty string, or invalid formats."""
    if val is None or val == "":
        return default
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return default


def safe_str(val: Any, default: str = "") -> str:
    """Safely convert any value to string, handling None."""
    if val is None:
        return default
    return str(val).strip()


class BaseFinanceExporter(ABC):
    """Abstract base class for reusable financial report exporters."""

    def __init__(self, save_path: str, college_name: str = "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"):
        if not save_path or not str(save_path).strip():
            raise ValueError("save_path must be a non-empty string path.")
        self.save_path = str(save_path).strip()
        self.college_name = (college_name or "วิทยาลัยอาชีวศึกษาสุพรรณบุรี").strip()

    @abstractmethod
    def export(self, data: Any, metadata: Dict[str, Any] = None) -> str:
        """Generates the output file and returns the saved path."""
        pass
