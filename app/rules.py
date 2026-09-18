from pathlib import Path
import json

DEFAULT_RULES = {
    "lead_time_import": 80,
    "lead_time_lokal": 15,
    "umur_berryman_import": 90,
    "umur_klevo_import": 90,
    "umur_berryman_lokal": 30,
    "ads_decimal_places": 2,
    "coverage_decimal_places": 1,
    "abnormal_multiplier": 3.0,
    "winning_safety_stock_pct": 20,
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
        "umur_berryman_import": max(1, int(rules["umur_berryman_import"])),
        "umur_klevo_import": max(1, int(rules["umur_klevo_import"])),
        "umur_berryman_lokal": max(1, int(rules["umur_berryman_lokal"])),
        "ads_decimal_places": min(6, max(0, int(rules["ads_decimal_places"]))),
        "coverage_decimal_places": min(6, max(0, int(rules["coverage_decimal_places"]))),
        "abnormal_multiplier": max(0.1, float(rules.get("abnormal_multiplier", 3.0))),
        "winning_safety_stock_pct": min(100, max(0, int(rules.get("winning_safety_stock_pct", 20)))),
    }
    (upload_dir / "purchasing_rules.json").write_text(json.dumps(normalized))
    return normalized
