import json
import re
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

from app.config import CUSTOMERS_FILE


def _load_raw() -> dict:
    if CUSTOMERS_FILE.exists():
        try:
            return json.loads(CUSTOMERS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def _save_raw(data: dict):
    CUSTOMERS_FILE.write_text(json.dumps(data, indent=2, default=str), encoding="utf-8")


def _normalize(name: str) -> str:
    return re.sub(r"\s+", " ", name.strip()).upper()


def load_customers() -> dict:
    return _load_raw()


def find_or_create_customer(raw_name: str) -> dict:
    name = raw_name.strip() if raw_name else ""
    if not name:
        name = "Unknown"
    key = _normalize(name)
    data = _load_raw()

    for cid, cust in data.items():
        if _normalize(cust.get("name", "")) == key:
            return {"id": cid, **cust}

    cid = str(uuid.uuid4())
    entry = {
        "name": name,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "reports": [],
    }
    data[cid] = entry
    _save_raw(data)
    return {"id": cid, **entry}


def add_report_to_customer(
    customer_id: str,
    report_type: str,
    original_filename: str,
    output_filename: str,
    period_label: Optional[str] = None,
    audit_filename: Optional[str] = None,
    net_capital_filename: Optional[str] = None,
):
    data = _load_raw()
    cust = data.get(customer_id)
    if not cust:
        return
    report = {
        "id": str(uuid.uuid4()),
        "type": report_type,
        "original_filename": original_filename,
        "output_filename": output_filename,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "period_label": period_label or "",
        "audit_filename": audit_filename or "",
        "net_capital_filename": net_capital_filename or "",
    }
    cust.setdefault("reports", []).append(report)
    _save_raw(data)


def get_customer(customer_id: str) -> Optional[dict]:
    data = _load_raw()
    cust = data.get(customer_id)
    if cust:
        return {"id": customer_id, **cust}
    return None
