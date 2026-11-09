from abc import ABC, abstractmethod
from typing import Any, Dict, List


class BaseFinanceExporter(ABC):
    """Abstract base class for reusable financial report exporters."""

    def __init__(self, save_path: str, college_name: str = "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"):
        self.save_path = save_path
        self.college_name = college_name or "วิทยาลัยอาชีวศึกษาสุพรรณบุรี"

    @abstractmethod
    def export(self, data: Any, metadata: Dict[str, Any] = None) -> str:
        """Generates the output file and returns the saved path."""
        pass
