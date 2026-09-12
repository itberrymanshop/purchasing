from pathlib import Path
from datetime import datetime
import re
import json
import shutil
import uuid
from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import HTMLResponse, FileResponse, StreamingResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import openpyxl
from .processor import process_workbook, save_report, load_active, load_audit
from .rules import load_rules, save_rules

BASE = Path(__file__).resolve().parent
ROOT = BASE.parent
UPLOADS = ROOT / "uploads"
UPLOADS.mkdir(exist_ok=True)
TEMPLATES = Jinja2Templates(directory=str(BASE / "templates"))
REQUIRED = {
    "STOK": ("stok", "ketersediaan stok penjualan"),
    "PENJUALAN_PERIODE": ("penjualan_periode", "penjualan barang per periode"),
    "PENJUALAN_TOTAL": ("penjualan_total", "penjualan per barang"),
    "MASTER_HARGA": ("master_harga", "master harga jual"),
    "DISCONTINUE": ("discontinue", "produk discontinue"),
    "IMPOR": ("impor", "barang masuk impor"),
    "PPN": ("ppn", "produk pm np"),
    "PARETO": ("pareto", "produk pareto"),
}

app = FastAPI(title="Purchasing Control")
app.mount("/static", StaticFiles(directory=str(BASE / "static")), name="static")


def merge_value(ws, row, column):
    for merged in ws.merged_cells.ranges:
        if merged.min_row <= row <= merged.max_row and merged.min_col <= column <= merged.max_col:
            return ws.cell(merged.min_row, merged.min_col).value
    return ws.cell(row, column).value


def normalized_headers(ws):
    max_col = ws.max_column
    top = [merge_value(ws, 5, col) for col in range(1, max_col + 1)]
    second = [merge_value(ws, 6, col) for col in range(1, max_col + 1)]
    result = []
    for first, child in zip(top, second):
        first_text = str(first or "").strip()
        child_text = str(child or "").strip()
        result.append(" - ".join(x for x in (first_text, child_text) if x))
    return result


def validate_workbook(path):
    workbook = openpyxl.load_workbook(path, read_only=False, data_only=True)
    names = workbook.sheetnames
    lower_names = {name.lower(): name for name in names}
    missing = []
    matched = {}
    for expected, aliases in REQUIRED.items():
        found = next((original for low, original in lower_names.items() if any(alias in low for alias in aliases)), None)
        if not found:
            missing.append(expected)
        else:
            matched[expected] = found
    sheets = []
    for expected, actual in matched.items():
        ws = workbook[actual]
        headers = normalized_headers(ws) if expected == "PENJUALAN_PERIODE" else [str(c.value or "").strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
        sheets.append({"name": actual, "expected": expected, "rows": max(0, ws.max_row - 1), "headers": headers[:18]})
    return missing, sheets


@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    latest = sorted(UPLOADS.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return TEMPLATES.TemplateResponse("index.html", {"request": request, "latest": latest[0].name if latest else None})


@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(request: Request):
    report = load_active(UPLOADS)
    if report and report.get("records") and "calculation" not in report["records"][0]:
        source = UPLOADS / report.get("filename", "")
        if source.exists():
            refreshed = process_workbook(source, load_rules(UPLOADS))
            report = {**report, "records": refreshed}
            (UPLOADS / "active_report.json").write_text(json.dumps(report, ensure_ascii=False))
    rows = report["records"] if report else []
    summary = {"total": len(rows), "critical": sum(row["status"] == "CRITICAL" for row in rows), "need_order": sum(row["status"] == "NEED ORDER" for row in rows), "sufficient": sum(row["status"] == "SUFFICIENT" for row in rows), "omzet": sum(row["total_sales"] for row in rows), "impor": sum(row["total_sales"] for row in rows if row["kind"] == "IMPOR"), "lokal": sum(row["total_sales"] for row in rows if row["kind"] == "LOKAL")}
    summary["impor_pct"] = summary["impor"] / summary["omzet"] * 100 if summary["omzet"] else 0
    summary["lokal_pct"] = summary["lokal"] / summary["omzet"] * 100 if summary["omzet"] else 0
    summary["pm"] = sum(row["total_sales"] for row in rows if row["tax"] == "PM")
    summary["np"] = sum(row["total_sales"] for row in rows if row["tax"] == "NP")
    summary["pm_pct"] = summary["pm"] / summary["omzet"] * 100 if summary["omzet"] else 0
    summary["np_pct"] = summary["np"] / summary["omzet"] * 100 if summary["omzet"] else 0
    short_months = ["JAN", "FEB", "MAR", "APR", "MEI", "JUN", "JUL", "AGT", "SEP", "OKT", "NOV", "DES"]
    month_order = rows[0]["month_order"] if rows else []
    month_labels = [{"number": month, "label": short_months[month - 1]} for month in month_order]
    projection_label = f"PROYEKSI {month_labels[-1]['label']}" if month_labels else "PROYEKSI"
    top_gainers = sorted((row for row in rows if row["change"] > 0), key=lambda row: row["change"], reverse=True)[:5]
    top_losers = sorted((row for row in rows if row["change"] < 0), key=lambda row: row["change"])[:5]
    top_revenue_desc = sorted(rows, key=lambda row: row["total_sales"], reverse=True)[:5]
    top_revenue_asc = sorted(rows, key=lambda row: row["total_sales"])[:5]
    trend_totals = [{"label": item["label"], "value": sum(row["months"].get(str(item["number"]), {}).get("subtotal", 0) for row in rows)} for item in month_labels]
    branch_trend = {branch: [{"label": item["label"], "value": sum(row["months"].get(str(item["number"]), {}).get(branch, 0) for row in rows)} for item in month_labels[:-1]] for branch in ("grosir", "marketplace", "klevo")}
    return TEMPLATES.TemplateResponse("dashboard.html", {"request": request, "rows": rows, "summary": summary, "report": report, "month_labels": month_labels, "projection_label": projection_label, "top_gainers": top_gainers, "top_losers": top_losers, "top_revenue_desc": top_revenue_desc, "top_revenue_asc": top_revenue_asc, "trend_totals": trend_totals, "branch_trend": branch_trend})


@app.get("/restock", response_class=HTMLResponse)
def restock(request: Request):
    report = load_active(UPLOADS)
    rows = [row for row in (report["records"] if report else []) if row.get("restock_qty", 0) > 0]
    return TEMPLATES.TemplateResponse("restock.html", {"request": request, "rows": rows, "report": report})

@app.get("/settings", response_class=HTMLResponse)
def settings(request: Request):
    return TEMPLATES.TemplateResponse("settings.html", {"request": request, "rules": load_rules(UPLOADS), "saved": request.query_params.get("saved")})

@app.post("/settings")
def save_settings(request: Request, lead_time_import: int = Form(...), lead_time_lokal: int = Form(...), safety_stock_percent: float = Form(...), minimum_order_quantity: int = Form(...), order_rounding: int = Form(...)):
    rules = save_rules(UPLOADS, locals())
    return RedirectResponse(url="/settings?saved=1", status_code=303)

@app.get("/audit", response_class=HTMLResponse)
def audit(request: Request):
    return TEMPLATES.TemplateResponse("audit.html", {"request": request, "reports": load_audit(UPLOADS)})


@app.get("/templates/download")
def download_template():
    from io import BytesIO
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    now = datetime.now()
    month_names = ["JAN", "FEB", "MAR", "APR", "MEI", "JUN", "JUL", "AGT", "SEP", "OKT", "NOV", "DES"]
    period_months = [((now.month - 4 + offset) % 12) + 1 for offset in range(4)]
    period_headers = []
    for month in period_months:
        label = f"{month_names[month - 1]} {now.year}"
        period_headers.extend([f"{label} - GROSIR", f"{label} - KLEVO", f"{label} - MARKETPLACE", f"{label} - Subtotal"])
    definitions = {
        "STOK": ["Nama Pemasok Utama", "Kode Barang", "Nama Barang", "Satuan", "CBM (cm3)", "Stok Gudang", "Dipesan", "Dijual", "Stok dapat dijual"],
        "PENJUALAN_PERIODE": ["Kode Barang", "Nama Barang", *period_headers],
        "PENJUALAN_TOTAL": ["Kode Barang", "Nama Barang", "Satuan", "Kuantitas", "Total Penjualan"],
        "MASTER_HARGA": ["Kode Barang", "Nama Barang", "Grosir Min"],
        "DISCONTINUE": ["Kode Barang", "Nama Barang", "Keterangan Barang"],
        "IMPOR": ["Kode Barang", "Nama Barang", "Keterangan"],
        "PPN": ["No.", "Pemasok Utama", "Kode Barang", "Nama Barang", "Ket. Barang", "JENIS", "PM / NP", "IMPOR"],
        "PARETO": ["Marketplace - No", "Marketplace - SKU", "Marketplace - Produk", "Grosir - No", "Grosir - SKU", "Grosir - Produk", "MP & GR - No", "MP & GR - SKU", "MP & GR - Produk"],
    }
    for name, headers in definitions.items():
        sheet = workbook.create_sheet(name)
        sheet.append(headers)
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(len(headers))}2"
        for cell in sheet[1]:
            cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
            cell.fill = openpyxl.styles.PatternFill("solid", fgColor="2563EB")
        for index, header in enumerate(headers, 1):
            sheet.column_dimensions[openpyxl.utils.get_column_letter(index)].width = min(34, max(14, len(header) + 3))
    guide = workbook.create_sheet("PETUNJUK")
    guide.append(["PETUNJUK PENGISIAN TEMPLATE PURCHASING"])
    for line in ["Jangan ubah nama sheet dan header.", "Copy-paste data 8 file sumber ke sheet masing-masing.", "Kode Barang harus berada di kolom A setiap sheet.", "Sheet penjualan periode memakai A sebagai SKU dan B:R sebagai empat grup bulan.", "Merge dari file Accurate boleh ikut terbawa; program menormalisasi nilainya.", "Simpan sebagai .xlsx lalu upload."]:
        guide.append([line])
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=template-purchasing.xlsx"})


@app.post("/upload", response_class=HTMLResponse)
async def upload(request: Request, file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xlsm")):
        return TEMPLATES.TemplateResponse("index.html", {"request": request, "error": "File harus berformat .xlsx atau .xlsm.", "latest": None}, status_code=422)
    safe_name = f"{datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:8]}.xlsx"
    path = UPLOADS / safe_name
    with path.open("wb") as output:
        shutil.copyfileobj(file.file, output)
    try:
        missing, sheets = validate_workbook(path)
    except Exception as exc:
        path.unlink(missing_ok=True)
        return TEMPLATES.TemplateResponse("index.html", {"request": request, "error": f"Excel tidak dapat dibaca: {exc}", "latest": None}, status_code=422)
    if missing:
        return TEMPLATES.TemplateResponse("index.html", {"request": request, "error": "Sheet wajib belum lengkap: " + ", ".join(missing), "validation": sheets, "latest": path.name}, status_code=422)
    try:
        records = process_workbook(path, load_rules(UPLOADS))
        report = save_report(UPLOADS, path.name, records)
    except Exception as exc:
        return TEMPLATES.TemplateResponse("index.html", {"request": request, "error": f"Proses laporan gagal: {exc}", "validation": sheets, "latest": path.name}, status_code=422)
    return RedirectResponse(url="/dashboard", status_code=303)
