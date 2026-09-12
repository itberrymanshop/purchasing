from pathlib import Path
import json

DEFAULT_RULES = {
    "lead_time_import": 80,
    "lead_time_lokal": 15,
    "safety_stock_percent": 20,
    "minimum_order_quantity": 1,
    "order_rounding": 1,
}

def load_rules(upload_dir: Path) -> dict:
    path = upload_dir / "purchasing_rules.json"
    if not path.exists():
        return DEFAULT_RULES.copy()
    try:
        saved = json.loads(path.read_text())
        return {**DEFAULT_RULES, **saved}
    except (OSError, json.JSONDecodeError):
        return DEFAULT_RULES.copy()

def save_rules(upload_dir: Path, rules: dict) -> dict:
    normalized = {
        "lead_time_import": max(1, int(rules["lead_time_import"])),
        "lead_time_lokal": max(1, int(rules["lead_time_lokal"])),
        "safety_stock_percent": max(0, float(rules["safety_stock_percent"])),
        "minimum_order_quantity": max(1, int(rules["minimum_order_quantity"])),
        "order_rounding": max(1, int(rules["order_rounding"])),
    }
    (upload_dir / "purchasing_rules.json").write_text(json.dumps(normalized))
    return normalized
