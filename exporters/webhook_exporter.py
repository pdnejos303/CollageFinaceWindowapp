import json
import threading
import urllib.request
import urllib.error
from typing import Any, Callable, Dict, Optional


class FinanceWebhookDispatcher:
    """Non-blocking Webhook Dispatcher for n8n / LINE Bot / Discord financial alerts (Reused from n8nXLineBot)."""

    def __init__(self, webhook_url: str, enabled: bool = False, alert_threshold: float = 50000.0):
        self.webhook_url = (webhook_url or "").strip()
        self.enabled = enabled and bool(self.webhook_url)
        self.alert_threshold = float(alert_threshold)

    def notify_event_async(
        self,
        event_type: str,
        title: str,
        amount: float,
        actor: str,
        ref_code: str = "",
        extra: Optional[Dict[str, Any]] = None,
        callback: Optional[Callable[[bool, str], None]] = None,
    ):
        if not self.enabled:
            return

        is_high_value = amount >= self.alert_threshold
        payload = {
            "app": "CollegeFinanceGodApp",
            "event": event_type,
            "title": title,
            "amount_thb": round(float(amount), 2),
            "is_high_value_alert": is_high_value,
            "actor": actor,
            "ref_code": ref_code,
            "details": extra or {},
        }

        thread = threading.Thread(
            target=self._worker,
            args=(payload, callback),
            daemon=True,
        )
        thread.start()

    def _worker(self, payload: Dict[str, Any], callback: Optional[Callable[[bool, str], None]]):
        try:
            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            req = urllib.request.Request(
                self.webhook_url,
                data=data,
                headers={
                    "Content-Type": "application/json; charset=utf-8",
                    "User-Agent": "CollegeFinanceGodApp/1.0",
                },
            )
            with urllib.request.urlopen(req, timeout=8) as resp:
                code = resp.getcode()
                ok = 200 <= code < 300
                if callback:
                    callback(ok, f"HTTP {code}")
        except urllib.error.URLError as e:
            if callback:
                callback(False, str(e))
        except Exception as e:
            if callback:
                callback(False, str(e))
