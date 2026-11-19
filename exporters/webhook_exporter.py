import json
import threading
import urllib.request
import urllib.error
from typing import Any, Callable, Dict, Optional
from exporters.base import safe_float, safe_str


class FinanceWebhookDispatcher:
    """Non-blocking Webhook Dispatcher for n8n / LINE Bot / Discord financial alerts (Reused from n8nXLineBot)."""

    def __init__(self, webhook_url: str, enabled: bool = False, alert_threshold: float = 50000.0):
        self.webhook_url = (webhook_url or "").strip()
        self.enabled = bool(enabled and self.webhook_url)
        self.alert_threshold = safe_float(alert_threshold, default=50000.0)

    def notify_event_async(
        self,
        event_type: str,
        title: str,
        amount: float,
        actor: str,
        ref_code: str = "",
        extra: Optional[Dict[str, Any]] = None,
        callback: Optional[Callable[[bool, str], None]] = None,
    ) -> Optional[threading.Thread]:
        """
        Dispatches financial notification event asynchronously in a background daemon thread.
        Returns the spawned Thread instance (or None if dispatcher is disabled).
        """
        if not self.enabled:
            return None

        clean_amount = safe_float(amount, default=0.0)
        is_high_value = clean_amount >= self.alert_threshold
        payload = {
            "app": "CollegeFinanceGodApp",
            "event": safe_str(event_type),
            "title": safe_str(title),
            "amount_thb": round(clean_amount, 2),
            "is_high_value_alert": is_high_value,
            "actor": safe_str(actor),
            "ref_code": safe_str(ref_code),
            "details": extra if isinstance(extra, dict) else {},
        }

        thread = threading.Thread(
            target=self._worker,
            args=(payload, callback),
            daemon=True,
        )
        thread.start()
        return thread

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
                    try:
                        callback(ok, f"HTTP {code}")
                    except Exception:
                        pass
        except urllib.error.HTTPError as e:
            if callback:
                try:
                    callback(False, f"HTTP {e.code}: {e.reason}")
                except Exception:
                    pass
        except urllib.error.URLError as e:
            if callback:
                try:
                    callback(False, str(e.reason if hasattr(e, 'reason') else e))
                except Exception:
                    pass
        except Exception as e:
            if callback:
                try:
                    callback(False, str(e))
                except Exception:
                    pass
