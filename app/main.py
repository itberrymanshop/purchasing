from pathlib import Path
from datetime import datetime
import re
import json
import shutil
import uuid
import os
from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import HTMLResponse, FileResponse, StreamingResponse, RedirectResponse, JSONResponse
from sqlalchemy import select
from sqlalchemy.orm import joinedload
from .auth import SessionLocal, User, Role, RoleProfile, Permission, Page, get_user, has_permission, initialize_auth, hash_password, verify_password, can_login_now, parse_time, DEFAULT_TIMEZONE
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import openpyxl
from io import BytesIO
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from .processor import process_workbook, save_report, load_active, load_audit
from .rules import load_rules, save_rules

BASE = Path(__file__).resolve().parent
ROOT = BASE.parent
UPLOADS = ROOT / "uploads"
UPLOADS.mkdir(exist_ok=True)
TEMPLATES = Jinja2Templates(directory=str(BASE / "templates"))
TEMPLATES.env.globals["has_permission"] = has_permission
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
initialize_auth()


@app.middleware("http")
async def load_user(request: Request, call_next):
    request.state.user = None
    user_id = request.cookies.get("user_id")
    with SessionLocal() as db:
        request.state.user = get_user(db, int(user_id)) if user_id and user_id.isdigit() else None
    return await call_next(request)


def require_user(request: Request, permission: str | None = None):
    user = request.state.user
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    if permission and not has_permission(user, permission):
        return TEMPLATES.TemplateResponse("forbidden.html", {"request": request}, status_code=403)
    return None


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


@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request):
    if request.state.user:
        return RedirectResponse(url="/dashboard", status_code=303)
    return TEMPLATES.TemplateResponse("login.html", {"request": request, "error": request.query_params.get("error")})


@app.post("/login")
def login(request: Request, username: str = Form(...), password: str = Form(...)):
    with SessionLocal() as db:
        user = db.scalar(select(User).where(User.username == username.strip().lower()))
        valid = user and user.is_active and verify_password(password, user.password_hash)
        if not valid:
            return RedirectResponse(url="/login?error=Kredensial tidak valid", status_code=303)
        if not can_login_now(user):
            return RedirectResponse(url="/login?error=Akun tidak diizinkan login pada jam ini", status_code=303)
        user.last_login_at = datetime.utcnow()
        db.commit()
        response = RedirectResponse(url="/dashboard", status_code=303)
        response.set_cookie("user_id", str(user.id), httponly=True, samesite="lax", secure=os.getenv("COOKIE_SECURE", "0") == "1", max_age=28800)
        return response


@app.post("/logout")
def logout():
    response = RedirectResponse(url="/login", status_code=303)
    response.delete_cookie("user_id")
    return response


@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    guard = require_user(request, "upload.create")
    if guard: return guard
    latest = sorted(UPLOADS.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return TEMPLATES.TemplateResponse("index.html", {"request": request, "latest": latest[0].name if latest else None})


@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(request: Request):
    guard = require_user(request, "dashboard.read")
    if guard: return guard
    report = load_active(UPLOADS)
    if report and report.get("records") and ("calculation" not in report["records"][0] or "inventory_age" not in report["records"][0]):
        source = UPLOADS / report.get("filename", "")
        if source.exists():
            refreshed = process_workbook(source, load_rules(UPLOADS))
            report = {**report, "records": refreshed}
            (UPLOADS / "active_report.json").write_text(json.dumps(report, ensure_ascii=False))
    rows = report["records"] if report else []
    display_rules = load_rules(UPLOADS)
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
    return TEMPLATES.TemplateResponse("dashboard.html", {"request": request, "rows": rows, "summary": summary, "report": report, "month_labels": month_labels, "projection_label": projection_label, "top_gainers": top_gainers, "top_losers": top_losers, "top_revenue_desc": top_revenue_desc, "top_revenue_asc": top_revenue_asc, "trend_totals": trend_totals, "branch_trend": branch_trend, "display_rules": display_rules})


@app.get("/restock", response_class=HTMLResponse)
def restock(request: Request):
    guard = require_user(request, "dashboard.read")
    if guard: return guard
    report = load_active(UPLOADS)
    rows = [row for row in (report["records"] if report else []) if row.get("status") in ("CRITICAL", "NEED ORDER")]
    return TEMPLATES.TemplateResponse("restock.html", {"request": request, "rows": rows, "report": report})

@app.get("/settings", response_class=HTMLResponse)
def settings(request: Request):
    guard = require_user(request, "settings.write")
    if guard: return guard
    return TEMPLATES.TemplateResponse("settings.html", {"request": request, "rules": load_rules(UPLOADS), "saved": request.query_params.get("saved")})

@app.get("/trends/{sku}")
def sku_trend(request: Request, sku: str):
    guard = require_user(request, "dashboard.read")
    if guard: return guard
    report = load_active(UPLOADS)
    rows = report["records"] if report else []
    row = next((item for item in rows if item["sku"].lower() == sku.lower()), None)
    if not row: return JSONResponse({"error": "SKU tidak ditemukan"}, status_code=404)
    labels = [{"number": month, "label": month} for month in row.get("month_order", [])]
    return JSONResponse({"sku": row["sku"], "name": row["name"], "labels": labels, "months": row.get("months", {})})


@app.post("/settings")
def save_settings(request: Request, lead_time_import: int = Form(...), lead_time_lokal: int = Form(...), umur_berryman_import: int = Form(...), umur_klevo_import: int = Form(...), umur_berryman_lokal: int = Form(...), ads_decimal_places: int = Form(...), coverage_decimal_places: int = Form(...)):
    guard = require_user(request, "settings.write")
    if guard: return guard
    rules = save_rules(UPLOADS, locals())
    return RedirectResponse(url="/settings?saved=1", status_code=303)

@app.get("/users", response_class=HTMLResponse)
def users(request: Request):
    guard = require_user(request, "users.manage")
    if guard: return guard
    with SessionLocal() as db:
        all_users = db.scalars(select(User).options(joinedload(User.roles)).order_by(User.username)).unique().all()
        roles = db.scalars(select(Role).options(joinedload(Role.permissions)).order_by(Role.name)).unique().all()
        permissions = db.scalars(select(Permission).order_by(Permission.code)).all()
        profiles = db.scalars(select(RoleProfile).options(joinedload(RoleProfile.roles)).order_by(RoleProfile.name)).unique().all()
        pages = db.scalars(select(Page).options(joinedload(Page.roles)).order_by(Page.kind, Page.label)).unique().all()
    return TEMPLATES.TemplateResponse("users.html", {"request": request, "users": all_users})


@app.get("/roles", response_class=HTMLResponse)
def roles_page(request: Request):
    guard = require_user(request, "users.manage")
    if guard: return guard
    with SessionLocal() as db:
        roles = db.scalars(select(Role).options(joinedload(Role.permissions)).order_by(Role.name)).unique().all()
        permissions = db.scalars(select(Permission).order_by(Permission.code)).all()
    return TEMPLATES.TemplateResponse("roles.html", {"request": request, "roles": roles, "permissions": permissions})


@app.get("/role-profiles", response_class=HTMLResponse)
def role_profiles_page(request: Request):
    guard = require_user(request, "users.manage")
    if guard: return guard
    with SessionLocal() as db:
        profiles = db.scalars(select(RoleProfile).options(joinedload(RoleProfile.roles)).order_by(RoleProfile.name)).unique().all()
        roles = db.scalars(select(Role).order_by(Role.name)).all()
    return TEMPLATES.TemplateResponse("role-profiles.html", {"request": request, "profiles": profiles, "roles": roles})


@app.get("/page-permissions", response_class=HTMLResponse)
def page_permissions_page(request: Request):
    guard = require_user(request, "users.manage")
    if guard: return guard
    with SessionLocal() as db:
        pages = db.scalars(select(Page).options(joinedload(Page.roles)).order_by(Page.kind, Page.label)).unique().all()
        roles = db.scalars(select(Role).order_by(Role.name)).all()
    return TEMPLATES.TemplateResponse("page-permissions.html", {"request": request, "pages": pages, "roles": roles})


@app.get("/users/new", response_class=HTMLResponse)
def new_user(request: Request):
    guard = require_user(request, "users.manage")
    if guard: return guard
    with SessionLocal() as db:
        roles = db.scalars(select(Role).order_by(Role.name)).all()
    return TEMPLATES.TemplateResponse("user-detail.html", {"request": request, "user": None, "roles": roles, "profiles": [], "default_timezone": DEFAULT_TIMEZONE})


@app.get("/users/{user_id}", response_class=HTMLResponse)
def user_detail(request: Request, user_id: int):
    guard = require_user(request, "users.manage")
    if guard: return guard
    with SessionLocal() as db:
        user = db.scalar(select(User).options(joinedload(User.roles)).where(User.id == user_id))
        roles = db.scalars(select(Role).order_by(Role.name)).all()
        profiles = db.scalars(select(RoleProfile).options(joinedload(RoleProfile.roles)).order_by(RoleProfile.name)).unique().all()
    if not user:
        return RedirectResponse(url="/users?error=User tidak ditemukan", status_code=303)
    return TEMPLATES.TemplateResponse("user-detail.html", {"request": request, "user": user, "roles": roles, "profiles": profiles, "default_timezone": DEFAULT_TIMEZONE})


@app.post("/roles")
def create_role(request: Request, role_name: str = Form(...), description: str = Form(""), permissions: list[str] = Form([])):
    guard = require_user(request, "users.manage")
    if guard: return guard
    with SessionLocal() as db:
        if db.scalar(select(Role).where(Role.name == role_name.strip())):
            return RedirectResponse(url="/users?error=Role sudah dipakai", status_code=303)
        permission_rows = db.scalars(select(Permission).where(Permission.code.in_(permissions))).all()
        db.add(Role(name=role_name.strip(), description=description.strip(), permissions=permission_rows))
        db.commit()
    return RedirectResponse(url="/users?saved=1", status_code=303)


@app.post("/page-permissions")
async def save_page_permissions(request: Request):
    guard = require_user(request, "users.manage")
    if guard: return guard
    form = await request.form()
    page_ids = [int(value) for value in form.getlist("page_ids")]
    with SessionLocal() as db:
        for page_id in page_ids:
            page = db.get(Page, page_id)
            if not page: continue
            role_ids = [int(value) for value in form.getlist(f"page_{page_id}_roles")]
            page.roles = db.scalars(select(Role).where(Role.id.in_(role_ids))).all() if role_ids else []
        db.commit()
    return RedirectResponse(url="/page-permissions?saved=1", status_code=303)


@app.post("/users/{user_id}/update")
def update_user(request: Request, user_id: int, full_name: str = Form(...), role_ids: list[int] = Form([]), role_profile_id: int | None = Form(None), timezone: str = Form(DEFAULT_TIMEZONE), login_after: str = Form(""), login_before: str = Form(""), is_active: bool = Form(False), new_password: str = Form("")):
    guard = require_user(request, "users.manage")
    if guard: return guard
    with SessionLocal() as db:
        user = db.get(User, user_id)
        if not user:
            return RedirectResponse(url="/users?error=User tidak ditemukan", status_code=303)
        roles = db.scalars(select(Role).where(Role.id.in_(role_ids))).all() if role_ids else []
        if role_profile_id:
            profile = db.scalar(select(RoleProfile).options(joinedload(RoleProfile.roles)).where(RoleProfile.id == role_profile_id))
            if profile:
                roles = profile.roles
        user.full_name = full_name.strip()
        user.roles = roles
        user.timezone = timezone or DEFAULT_TIMEZONE
        user.login_after = parse_time(login_after)
        user.login_before = parse_time(login_before)
        user.is_active = is_active
        if new_password:
            user.password_hash = hash_password(new_password)
        db.commit()
    return RedirectResponse(url="/users?saved=1", status_code=303)


@app.post("/role-profiles")
def create_role_profile(request: Request, profile_name: str = Form(...), profile_description: str = Form(""), profile_role_ids: list[int] = Form([])):
    guard = require_user(request, "users.manage")
    if guard: return guard
    with SessionLocal() as db:
        if db.scalar(select(RoleProfile).where(RoleProfile.name == profile_name.strip())):
            return RedirectResponse(url="/users?error=Role profile sudah dipakai", status_code=303)
        selected_roles = db.scalars(select(Role).where(Role.id.in_(profile_role_ids))).all() if profile_role_ids else []
        db.add(RoleProfile(name=profile_name.strip(), description=profile_description.strip(), roles=selected_roles))
        db.commit()
    return RedirectResponse(url="/users?saved=1", status_code=303)


@app.post("/users", response_class=HTMLResponse)
def create_user(request: Request, username: str = Form(...), full_name: str = Form(...), password: str = Form(...), role_ids: list[int] = Form([]), timezone: str = Form(DEFAULT_TIMEZONE), login_after: str = Form(""), login_before: str = Form("")):
    guard = require_user(request, "users.manage")
    if guard: return guard
    with SessionLocal() as db:
        if db.scalar(select(User).where(User.username == username.strip().lower())):
            return RedirectResponse(url="/users?error=Username sudah dipakai", status_code=303)
        selected_roles = db.scalars(select(Role).where(Role.id.in_(role_ids))).all() if role_ids else []
        if not selected_roles or not password:
            return RedirectResponse(url="/users?error=Data user tidak valid", status_code=303)
        db.add(User(username=username.strip().lower(), full_name=full_name.strip(), password_hash=hash_password(password), roles=selected_roles, timezone=timezone or DEFAULT_TIMEZONE, login_after=parse_time(login_after), login_before=parse_time(login_before)))
        db.commit()
    return RedirectResponse(url="/users?saved=1", status_code=303)


@app.get("/audit", response_class=HTMLResponse)
def audit(request: Request):
    guard = require_user(request, "audit.read")
    if guard: return guard
    return TEMPLATES.TemplateResponse("audit.html", {"request": request, "reports": load_audit(UPLOADS)})


def export_rows():
    report = load_active(UPLOADS)
    return report, report["records"] if report else [], load_rules(UPLOADS)


TITLE_FONT = Font(name="Calibri", size=15, bold=True, color="17231E")
SUBTITLE_FONT = Font(name="Calibri", size=9, color="68756D")
HEADER_FONT = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
NOTE_FONT = Font(name="Calibri", size=9, italic=True, color="1D4F35")
BODY_FONT = Font(name="Calibri", size=9)
THIN_BORDER = Border(left=Side(style="thin", color="DCE4DD"), right=Side(style="thin", color="DCE4DD"), top=Side(style="thin", color="DCE4DD"), bottom=Side(style="thin", color="DCE4DD"))
STATUS_LABELS = {"CRITICAL": "Kritis", "NEED ORDER": "Order", "SUFFICIENT": "Aman", "DISCONTINUE": "DISCONTINUE"}
STATUS_FILLS = {"CRITICAL": PatternFill("solid", fgColor="FBE7E5"), "Kritis": PatternFill("solid", fgColor="FBE7E5"), "NEED ORDER": PatternFill("solid", fgColor="FFF0DC"), "Order": PatternFill("solid", fgColor="FFF0DC"), "DISCONTINUE": PatternFill("solid", fgColor="F3DAD8"), "SUFFICIENT": PatternFill("solid", fgColor="E6F5E8"), "Aman": PatternFill("solid", fgColor="E6F5E8")}


def format_export_sheet(ws, title, subtitle, headers, status_column=None):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A4"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    title_cell = ws.cell(row=1, column=1, value=title)
    title_cell.font = TITLE_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    subtitle_cell = ws.cell(row=2, column=1, value=subtitle)
    subtitle_cell.font = SUBTITLE_FONT
    ws.append(headers)
    for cell in ws[3]:
        cell.font = HEADER_FONT
        cell.fill = PatternFill("solid", fgColor="286645")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    ws.row_dimensions[3].height = 30
    ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}3"
    return status_column


def finalize_export_sheet(ws, status_column=None, money_columns=(), integer_columns=(), decimal_columns=()):
    status_index = next((cell.column - 1 for cell in ws[3] if cell.value == status_column), None) if status_column else None
    for row in ws.iter_rows(min_row=4):
        zebra = PatternFill("solid", fgColor="F8FAF8") if row[0].row % 2 == 0 else PatternFill()
        for cell in row:
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.fill = zebra
        if status_index is not None:
            status = str(row[status_index].value or "").upper()
            highlight = STATUS_FILLS.get(status)
            if highlight:
                for cell in row:
                    cell.fill = highlight
        for index in money_columns:
            if isinstance(row[index].value, (int, float)):
                row[index].number_format = '#,##0'
                row[index].alignment = Alignment(horizontal="right", vertical="center")
        for index in integer_columns:
            if isinstance(row[index].value, (int, float)):
                row[index].number_format = '#,##0'
                row[index].alignment = Alignment(horizontal="right", vertical="center")
        for index in decimal_columns:
            if isinstance(row[index].value, (int, float)):
                row[index].number_format = '0.00'
                row[index].alignment = Alignment(horizontal="right", vertical="center")
            elif isinstance(row[index].value, str) and row[index].value.startswith("="):
                row[index].number_format = '0.00'
                row[index].alignment = Alignment(horizontal="right", vertical="center")
    for column_index in range(1, ws.max_column + 1):
        width = 8
        for row in ws.iter_rows(min_row=4, min_col=column_index, max_col=column_index):
            width = max(width, len(str(row[0].value or "")))
        ws.column_dimensions[get_column_letter(column_index)].width = min(46, max(12, width + 3))
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


@app.get("/export/excel")
def export_excel(request: Request):
    guard = require_user(request, "report.export")
    if guard: return guard
    report, rows, rules = export_rows()
    if not report: return RedirectResponse(url="/dashboard", status_code=303)
    generated_at = datetime.now().strftime("%d %b %Y %H:%M")
    source = f"Sumber: {report['filename']} · Diproses: {report['processed_at']} · Diunduh: {generated_at}"
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    coverage = wb.create_sheet("ANALISIS COVERAGE")
    short_months = ["JAN","FEB","MAR","APR","MEI","JUN","JUL","AGT","SEP","OKT","NOV","DES"]
    coverage_month = short_months[rows[0]["month_order"][-1] - 1] if rows else "-"
    coverage_days = rows[0]["calculation"]["sales_days"] if rows else 0
    coverage_month_headers = [f"{short_months[month - 1]}" for month in rows[0]["month_order"][:3]] if rows else ["Bulan -3", "Bulan -2", "Bulan -1"]
    coverage_headers = ["SKU", "Nama Barang", *coverage_month_headers, "Proyeksi Bulan Berjalan", "Rata-rata 3 Bulan", "ADS", "Coverage On Hand", "Coverage + OTW", "Status"]
    format_export_sheet(coverage, "ANALISIS COVERAGE", f"Bulan berjalan {coverage_month} ({coverage_days} hari) · ADS memakai rata-rata 3 bulan histori ditambah proyeksi bulan berjalan · {source}", coverage_headers, status_column="Status")
    for i,row in enumerate(rows,4):
        months=row["month_order"][:3]; coverage.append([row["sku"],row["name"],*[row["months"].get(str(m),{}).get("subtotal",0) for m in months],row["projection"],f"=AVERAGE(C{i}:E{i})",f"=(G{i}+F{i})/{coverage_days}",f'=IF(H{i}=0,"",{row["available"]}/H{i})',f'=IF(H{i}=0,"",{row["available"]+row["ordered"]}/H{i})',STATUS_LABELS.get(row["status"], row["status"])])
    coverage_status = len(coverage_headers) - 1
    finalize_export_sheet(coverage, status_column=coverage_headers[coverage_status], integer_columns=(2,3,4,5,6), decimal_columns=(7,8,9))
    restock_rows = [row for row in rows if row["status"] in ("CRITICAL", "NEED ORDER")]
    restock = wb.create_sheet("PURCHASING ACTION")
    restock_headers = ["SKU","Nama Barang","Jenis","Umur Persediaan","Coverage + OTW","Stok + OTW","Rekomendasi Stok","Alasan","Status"]
    format_export_sheet(restock, "PURCHASING ACTION", f"Rekomendasi stok · {source}", restock_headers, status_column="Status")
    for row in restock_rows:
        restock.append([row["sku"],row["name"],row["kind"],row["inventory_age"],row["on_way"],row["available"]+row["ordered"],row["target_stock"],row["restock_reason"],STATUS_LABELS.get(row["status"], row["status"])])
    finalize_export_sheet(restock, status_column=restock_headers[-1], integer_columns=(1,2,3,5,6))
    detail_headers = ["No.","SKU","Nama Barang","Trend Bulan -3","Trend Bulan -2","Trend Bulan -1","Supplier","CBM","Stok Gudang","Dipesan","Dijual","Stok Dapat Dijual","Ket. Barang","Jenis","PM / NP","Keterangan Impor"]
    for month in rows[0]["month_order"] if rows else []: detail_headers += [f"{month} Grosir",f"{month} Marketplace",f"{month} Klevo",f"{month} Subtotal"]
    detail_headers += ["Stok","Bulan Lalu","Proyeksi","Coverage","OTW","Change","Omzet","Rekomendasi Stok","Alasan","Status"]
    detail = wb.create_sheet("DETAIL STOK DAN PENJUALAN")
    format_export_sheet(detail, "DETAIL STOK DAN PENJUALAN", f"Kolom detail mengikuti tabel Detail stok dan penjualan di Dashboard · {source}", detail_headers, status_column="Status")
    for row in rows:
        trend_months = row["month_order"][:3]
        trend_values = [row["months"].get(str(month), {}).get("subtotal", 0) for month in trend_months]
        values=[row["no"],row["sku"],row["name"],*trend_values,row["supplier"],row["cbm"],row["stock_gudang"],row["ordered"],row["sold"],row["available"],row["condition"],row["kind"],row["tax"],row["import_detail"]]
        for month in row["month_order"]:
            data=row["months"].get(str(month),{}); values += [data.get("grosir",0),data.get("marketplace",0),data.get("klevo",0),data.get("subtotal",0)]
        values += [row["available"],row["previous_sales"],row["projection"],round(row["on_hand"],rules["coverage_decimal_places"]) if row["on_hand"] is not None else None,round(row["on_way"],rules["coverage_decimal_places"]) if row["on_way"] is not None else None,row["change"],row["total_sales"],row["target_stock"],row["restock_reason"],STATUS_LABELS.get(row["status"], row["status"])]
        detail.append(values)
    finalize_export_sheet(detail, status_column=detail_headers[-1])
    for ws, notes in [(coverage,[("RUMUS", "Rata-rata 3 Bulan = AVERAGE(Bulan -3:Bulan -1); ADS = (Rata-rata 3 Bulan + Proyeksi Bulan Berjalan) / Hari Bulan Berjalan; Coverage On Hand = Available / ADS; Coverage + OTW = (Available + Dipesan) / ADS")]),(restock,[("RUMUS", "Rekomendasi Stok = ceil(ADS × Umur Persediaan). Umur persediaan: 90 hari untuk Berryman impor dan Klevo impor, 30 hari untuk Berryman lokal.")]),(detail,[("KETERANGAN", "Kolom detail mengikuti tabel Detail stok dan penjualan di Dashboard. Nilai coverage dan OTW dibulatkan hanya untuk tampilan/export.")])]:
        note_row = ws.max_row + 2
        ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=ws.max_column)
        ws.cell(row=note_row, column=1, value=f"{notes[0][0]}: {notes[0][1]}").font = NOTE_FONT
    output=BytesIO(); wb.save(output); output.seek(0)
    return StreamingResponse(output,media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",headers={"Content-Disposition":f'attachment; filename="laporan-purchasing-{datetime.now():%Y%m%d-%H%M}.xlsx"'})


@app.get("/export/pdf")
def export_pdf(request: Request):
    guard = require_user(request, "report.export")
    if guard: return guard
    if not load_active(UPLOADS): return RedirectResponse(url="/dashboard", status_code=303)
    return RedirectResponse(url="/dashboard?print=1", status_code=303)


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
    guard = require_user(request, "upload.create")
    if guard: return guard
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
