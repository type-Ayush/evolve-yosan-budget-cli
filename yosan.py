import argparse
import calendar
from datetime import datetime
import json
import os
from pathlib import Path
import sqlite3
import sys
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import HRFlowable, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from auth import clear_session, delete_account_flow, require_login, show_profile_view

# Enable Windows ANSI Virtual Terminal Processing
os.system("")

DATA_DIR = Path.home() / "Documents"
REPORTS_DIR = DATA_DIR / "yosan_reports"


# ==========================================
# 🎨 COLOR CONSTANTS
# ==========================================
class C:
    RESET   = "\033[0m"
    BOLD    = "\033[1m"
    DIM     = "\033[2m"
    CYAN    = "\033[38;2;0;210;255m"
    GREEN   = "\033[38;2;46;204;113m"
    RED     = "\033[38;2;231;76;60m"
    YELLOW  = "\033[38;2;241;196;15m"
    PURPLE  = "\033[38;2;165;94;234m"
    ORANGE  = "\033[38;2;250;130;49m"
    GRAY    = "\033[38;2;127;143;166m"
    WHITE   = "\033[38;2;245;246;250m"


CATEGORIES = {
    "-m": "Mess Food",
    "-c": "Clothes",
    "-a": "Accessories",
    "-v": "Savings",
}

CATEGORY_ANSI = {
    "Mess Food": "\033[38;2;52;152;219m",    # Vibrant Blue
    "Clothes": "\033[38;2;46;204;113m",      # Emerald Green
    "Accessories": "\033[38;2;243;156;18m",  # Amber / Gold
    "Savings": "\033[38;2;155;89;182m",      # Royal Purple
}

CATEGORY_COLORS = {
    "Mess Food": "2E75B6",
    "Clothes": "548235",
    "Accessories": "BF8F00",
    "Savings": "7030A0",
}


# ==========================================
# 1. DATABASE & USER PATH ISOLATION
# ==========================================
def get_current_username() -> str:
    session_file = DATA_DIR / ".yosan_session.json"
    if session_file.exists():
        try:
            with open(session_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                uname = data.get("username")
                if uname and uname.strip():
                    return uname.strip()
        except Exception:
            pass
    return "default"


def get_user_db_path() -> str:
    username = get_current_username()
    safe_name = "".join(c for c in username if c.isalnum() or c in ("-", "_")).lower()
    return str(DATA_DIR / f"yosan_{safe_name}.db")


def get_user_excel_path() -> str:
    username = get_current_username()
    safe_name = "".join(c for c in username if c.isalnum() or c in ("-", "_")).lower()
    return str(DATA_DIR / f"budget_book_{safe_name}.xlsx")


def get_db_connection():
    db_file = get_user_db_path()
    conn = sqlite3.connect(db_file)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS months (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                month_code TEXT UNIQUE NOT NULL,
                month_name TEXT NOT NULL,
                total_budget REAL NOT NULL,
                is_active INTEGER DEFAULT 1,
                created_at TEXT NOT NULL
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS allocations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                month_id INTEGER NOT NULL,
                branch TEXT NOT NULL,
                allocated_amount REAL NOT NULL,
                base_amount REAL DEFAULT 0.0,
                FOREIGN KEY (month_id) REFERENCES months(id) ON DELETE CASCADE,
                UNIQUE(month_id, branch)
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                month_id INTEGER NOT NULL,
                branch TEXT NOT NULL,
                description TEXT NOT NULL,
                amount REAL NOT NULL,
                timestamp TEXT NOT NULL,
                FOREIGN KEY (month_id) REFERENCES months(id) ON DELETE CASCADE
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS topups (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                month_id INTEGER NOT NULL,
                branch TEXT NOT NULL,
                amount REAL NOT NULL,
                description TEXT DEFAULT 'Budget Top-up',
                timestamp TEXT NOT NULL,
                FOREIGN KEY (month_id) REFERENCES months(id) ON DELETE CASCADE
            )
        """)

        cursor.execute("PRAGMA table_info(allocations)")
        columns_alloc = [row["name"] for row in cursor.fetchall()]
        if "base_amount" not in columns_alloc:
            cursor.execute("ALTER TABLE allocations ADD COLUMN base_amount REAL DEFAULT 0.0")
            cursor.execute("UPDATE allocations SET base_amount = allocated_amount WHERE base_amount = 0.0 OR base_amount IS NULL")

        cursor.execute("PRAGMA table_info(topups)")
        columns_topup = [row["name"] for row in cursor.fetchall()]
        if "description" not in columns_topup:
            cursor.execute("ALTER TABLE topups ADD COLUMN description TEXT DEFAULT 'Budget Top-up'")

        conn.commit()


def get_active_month():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM months WHERE is_active = 1 LIMIT 1")
        return cursor.fetchone()


# ==========================================
# 2. EXCEL SYNC
# ==========================================
def create_borders(thin_color="D9D9D9"):
    return Border(
        left=Side(style="thin", color=thin_color),
        right=Side(style="thin", color=thin_color),
        top=Side(style="thin", color=thin_color),
        bottom=Side(style="thin", color=thin_color),
    )


def sync_to_excel():
    active_m = get_active_month()
    if not active_m:
        return

    wb = openpyxl.Workbook()
    ws_dash = wb.active
    ws_dash.title = "Summary"
    ws_dash.append([
        "Category",
        "Base Budget (₹)",
        "Credited (₹)",
        "Total Allocated (₹)",
        "Spent (₹)",
        "Remaining (₹)",
    ])

    dash_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    dash_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    for cell in ws_dash[1]:
        cell.fill = dash_fill
        cell.font = dash_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = create_borders()

    with get_db_connection() as conn:
        cursor = conn.cursor()
        row_idx = 2
        grand_base = 0.0
        grand_credit = 0.0
        grand_alloc = 0.0
        grand_spent = 0.0

        for cat_name in CATEGORIES.values():
            cursor.execute(
                "SELECT allocated_amount, base_amount FROM allocations WHERE month_id = ? AND branch = ?",
                (active_m["id"], cat_name),
            )
            alloc_row = cursor.fetchone()
            alloc_amt = alloc_row["allocated_amount"] if alloc_row else 0.0
            base_amt = (
                alloc_row["base_amount"]
                if (alloc_row and alloc_row["base_amount"] is not None and alloc_row["base_amount"] > 0)
                else alloc_amt
            )

            cursor.execute(
                "SELECT SUM(amount) AS total FROM topups WHERE month_id = ? AND branch = ?",
                (active_m["id"], cat_name),
            )
            topup_row = cursor.fetchone()
            credited_amt = topup_row["total"] if topup_row["total"] else 0.0

            cursor.execute(
                "SELECT SUM(amount) AS total FROM transactions WHERE month_id = ? AND branch = ?",
                (active_m["id"], cat_name),
            )
            spent_row = cursor.fetchone()
            spent_amt = spent_row["total"] if spent_row["total"] else 0.0

            remaining = alloc_amt - spent_amt
            grand_base += base_amt
            grand_credit += credited_amt
            grand_alloc += alloc_amt
            grand_spent += spent_amt

            ws_dash.append([
                cat_name,
                base_amt,
                credited_amt,
                alloc_amt,
                spent_amt,
                remaining,
            ])
            for c in ws_dash[row_idx]:
                c.font = Font(name="Calibri", size=10)
                c.border = create_borders("E0E0E0")
            row_idx += 1

            # Branch Tab
            ws_cat = wb.create_sheet(title=cat_name)
            ws_cat.append(["Timestamp", "Type", "Description", "Amount (₹)"])
            color = CATEGORY_COLORS[cat_name]
            h_fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            for cell in ws_cat[1]:
                cell.fill = h_fill
                cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = create_borders()
            ws_cat.row_dimensions[1].height = 24

            cursor.execute(
                """
                SELECT timestamp, 'Expense' AS type, description, amount FROM transactions WHERE month_id = ? AND branch = ?
                UNION ALL
                SELECT timestamp, 'Credit' AS type, description, amount FROM topups WHERE month_id = ? AND branch = ?
                ORDER BY timestamp ASC
                """,
                (active_m["id"], cat_name, active_m["id"], cat_name),
            )
            for tx in cursor.fetchall():
                ws_cat.append([
                    tx["timestamp"],
                    tx["type"],
                    tx["description"],
                    tx["amount"],
                ])
                last_r = ws_cat.max_row
                ws_cat.cell(row=last_r, column=1).alignment = Alignment(horizontal="center")
                ws_cat.cell(row=last_r, column=2).alignment = Alignment(horizontal="center")
                ws_cat.cell(row=last_r, column=4).alignment = Alignment(horizontal="right")
                for cell in ws_cat[last_r]:
                    cell.border = create_borders("E0E0E0")

            for col in ws_cat.columns:
                max_l = max(len(str(cell.value or "")) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws_cat.column_dimensions[col_letter].width = max(max_l + 5, 14)

        ws_dash.append([
            "TOTAL",
            grand_base,
            grand_credit,
            grand_alloc,
            grand_spent,
            grand_alloc - grand_spent,
        ])
        for c in ws_dash[row_idx]:
            c.font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
            c.border = create_borders()

        for col in ws_dash.columns:
            max_l = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws_dash.column_dimensions[col_letter].width = max(max_l + 5, 14)

    target_excel_path = get_user_excel_path()
    try:
        wb.save(target_excel_path)
    except PermissionError:
        pass


# ==========================================
# 3. REPORT EXPORTERS (TXT & PDF)
# ==========================================
def generate_txt_report(target_m, branch_data, grand_base, grand_credit, grand_alloc, grand_spent):
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    username = get_current_username()
    file_path = REPORTS_DIR / f"budget_report_{username}_{target_m['month_code']}.txt"

    header_title = (
        f"MONTH-END STATEMENT: {target_m['month_name'].upper()}"
        if target_m["is_active"] == 0
        else f"INTERIM STATEMENT (IN PROGRESS): {target_m['month_name'].upper()}"
    )

    lines = [
        "=" * 98,
        f"                {header_title} [{target_m['month_code']}]",
        "=" * 98,
        f" User           : {username}",
        f" Status         : {'BURNED IN / ARCHIVED (FINAL)' if target_m['is_active'] == 0 else 'ACTIVE / ONGOING'}",
        f" Generated on   : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f" Base Budget    : INR {grand_base:.2f}",
        f" Credited (+)   : INR {grand_credit:.2f}",
        f" Total Budget   : INR {grand_alloc:.2f}",
        f" Total Spent    : INR {grand_spent:.2f}",
        f" Net Balance    : INR {(grand_alloc - grand_spent):.2f}",
        "=" * 98 + "\n",
    ]

    for cat_name, b_info in branch_data.items():
        base = b_info["base"]
        credit = b_info["credit"]
        alloc = b_info["alloc"]
        spent = b_info["spent"]
        rem = alloc - spent
        lines.append(f"BRANCH: {cat_name.upper()}")
        lines.append(f"Base: INR {base:.2f} | Credited: +INR {credit:.2f} | Total: INR {alloc:.2f} | Spent: INR {spent:.2f} | Remaining: INR {rem:.2f}")
        lines.append("-" * 98)
        lines.append(f"{'#':<4} | {'Timestamp':<20} | {'Description':<38} | {'Debit (INR)':>13} | {'Credit (INR)':>13}")
        lines.append("-" * 98)
        if b_info["ledger"]:
            for idx, tx in enumerate(b_info["ledger"], start=1):
                debit_str = f"{tx['amount']:>13.2f}" if tx["type"] == "Expense" else f"{'-':>13}"
                credit_str = f"{tx['amount']:>13.2f}" if tx["type"] == "Credit" else f"{'-':>13}"
                lines.append(f"{idx:<4} | {tx['timestamp']:<20} | {tx['description'][:38]:<38} | {debit_str} | {credit_str}")
        else:
            lines.append("     No transactions or top-ups recorded for this branch.")
        lines.append("-" * 98 + "\n")

    with open(file_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    return file_path


def generate_pdf_report(target_m, branch_data, grand_base, grand_credit, grand_alloc, grand_spent):
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    username = get_current_username()
    file_path = REPORTS_DIR / f"budget_report_{username}_{target_m['month_code']}.pdf"

    doc = SimpleDocTemplate(
        str(file_path),
        pagesize=A4,
        rightMargin=32,
        leftMargin=32,
        topMargin=32,
        bottomMargin=32,
    )
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "ReportTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=22,
        textColor=colors.black,
    )
    meta_style = ParagraphStyle(
        "ReportMeta",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.black,
    )
    section_head = ParagraphStyle(
        "SectionHead",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        textColor=colors.black,
    )
    cell_bold_white = ParagraphStyle(
        "CellBoldWhite",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=10,
        textColor=colors.white,
    )
    cell_bold = ParagraphStyle(
        "CellBold",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=10,
    )
    cell_regular = ParagraphStyle(
        "CellRegular",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=10,
    )

    if target_m["is_active"] == 0:
        main_header = f"MONTH-END STATEMENT: {target_m['month_name'].upper()}"
        status_label = "BURNED IN / ARCHIVED (FINAL STATEMENT)"
    else:
        main_header = f"INTERIM STATEMENT: {target_m['month_name'].upper()}"
        status_label = "ACTIVE BUDGET CYCLE (IN PROGRESS)"

    story.append(Paragraph(main_header, title_style))
    story.append(Spacer(1, 4))
    story.append(
        Paragraph(
            f"User: <b>{username}</b> &nbsp;|&nbsp; Code: <b>{target_m['month_code']}</b> &nbsp;|&nbsp; Status: <b>{status_label}</b> &nbsp;|&nbsp; Generated: <b>{datetime.now().strftime('%d %b %Y, %H:%M')}</b>",
            meta_style,
        )
    )
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.black, spaceAfter=14))

    summary_data = [
        [
            Paragraph("Category", cell_bold_white),
            Paragraph("Base Alloc", cell_bold_white),
            Paragraph("Credited (+)", cell_bold_white),
            Paragraph("Spent Amount", cell_bold_white),
            Paragraph("Remaining Balance", cell_bold_white),
        ]
    ]

    for cat_name, b_info in branch_data.items():
        rem = b_info["alloc"] - b_info["spent"]
        summary_data.append([
            Paragraph(cat_name, cell_regular),
            f"INR {b_info['base']:.2f}",
            f"+INR {b_info['credit']:.2f}" if b_info["credit"] > 0 else "INR 0.00",
            f"INR {b_info['spent']:.2f}",
            f"INR {rem:.2f}",
        ])

    summary_data.append([
        Paragraph("TOTAL", cell_bold),
        f"INR {grand_base:.2f}",
        f"+INR {grand_credit:.2f}",
        f"INR {grand_spent:.2f}",
        f"INR {(grand_alloc - grand_spent):.2f}",
    ])

    t_summary = Table(summary_data, colWidths=[135, 95, 95, 100, 105])
    t_summary.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.black),
            ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ("LINEBELOW", (0, -1), (-1, -1), 1.5, colors.black),
        ])
    )

    story.append(Paragraph("EXECUTIVE SUMMARY", section_head))
    story.append(Spacer(1, 6))
    story.append(t_summary)
    story.append(Spacer(1, 18))

    story.append(Paragraph("DETAILED TRANSACTION LEDGER", section_head))
    story.append(Spacer(1, 8))

    for cat_name, b_info in branch_data.items():
        alloc = b_info["alloc"]
        spent = b_info["spent"]
        credit = b_info["credit"]
        rem = alloc - spent

        credit_str = f" | Credited: +INR {credit:.2f}" if credit > 0 else ""
        story.append(
            Paragraph(
                f"<b>{cat_name.upper()}</b> &nbsp;—&nbsp; Effective Budget: INR {alloc:.2f}{credit_str} | Spent: INR {spent:.2f} | Balance: INR {rem:.2f}",
                meta_style,
            )
        )
        story.append(Spacer(1, 4))

        branch_table_data = [
            [
                Paragraph("#", cell_bold),
                Paragraph("Timestamp", cell_bold),
                Paragraph("Description", cell_bold),
                Paragraph("Debit (INR)", cell_bold),
                Paragraph("Credit (INR)", cell_bold),
            ]
        ]

        if b_info["ledger"]:
            for idx, tx in enumerate(b_info["ledger"], start=1):
                debit_val = f"INR {tx['amount']:.2f}" if tx["type"] == "Expense" else "-"
                credit_val = f"+INR {tx['amount']:.2f}" if tx["type"] == "Credit" else "-"
                branch_table_data.append([
                    Paragraph(str(idx), cell_regular),
                    Paragraph(tx["timestamp"], cell_regular),
                    Paragraph(tx["description"], cell_regular),
                    debit_val,
                    credit_val,
                ])
        else:
            branch_table_data.append([
                Paragraph("-", cell_regular),
                Paragraph("No records", cell_regular),
                Paragraph("-", cell_regular),
                "-",
                "-",
            ])

        t_branch = Table(branch_table_data, colWidths=[25, 125, 190, 95, 95])
        t_branch.setStyle(
            TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAEAEA")),
                ("ALIGN", (3, 0), (4, -1), "RIGHT"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ])
        )
        story.append(t_branch)
        story.append(Spacer(1, 12))

    doc.build(story)
    return file_path


def export_monthly_report(month_code: str = None):
    init_db()
    with get_db_connection() as conn:
        cursor = conn.cursor()
        if month_code:
            cursor.execute("SELECT * FROM months WHERE month_code = ?", (month_code,))
            target_m = cursor.fetchone()
        else:
            active_m = get_active_month()
            if active_m:
                target_m = active_m
            else:
                cursor.execute("SELECT * FROM months ORDER BY id DESC LIMIT 1")
                target_m = cursor.fetchone()

    if not target_m:
        print(f"\n{C.RED}❌ No budget data found to generate reports.{C.RESET}")
        return

    month_id = target_m["id"]
    branch_data = {}
    grand_base = 0.0
    grand_credit = 0.0
    grand_alloc = 0.0
    grand_spent = 0.0

    with get_db_connection() as conn:
        cursor = conn.cursor()
        for cat_name in CATEGORIES.values():
            cursor.execute(
                "SELECT allocated_amount, base_amount FROM allocations WHERE month_id = ? AND branch = ?",
                (month_id, cat_name),
            )
            alloc_row = cursor.fetchone()
            alloc_amt = alloc_row["allocated_amount"] if alloc_row else 0.0
            base_amt = (
                alloc_row["base_amount"]
                if (alloc_row and alloc_row["base_amount"] is not None and alloc_row["base_amount"] > 0)
                else alloc_amt
            )

            cursor.execute(
                "SELECT SUM(amount) AS total FROM topups WHERE month_id = ? AND branch = ?",
                (month_id, cat_name),
            )
            topup_row = cursor.fetchone()
            credit_amt = topup_row["total"] if topup_row["total"] else 0.0

            cursor.execute(
                "SELECT SUM(amount) AS total FROM transactions WHERE month_id = ? AND branch = ?",
                (month_id, cat_name),
            )
            spent_row = cursor.fetchone()
            cat_spent = spent_row["total"] if spent_row["total"] else 0.0

            cursor.execute(
                """
                SELECT timestamp, 'Expense' AS type, description, amount FROM transactions WHERE month_id = ? AND branch = ?
                UNION ALL
                SELECT timestamp, 'Credit' AS type, description, amount FROM topups WHERE month_id = ? AND branch = ?
                ORDER BY timestamp ASC
                """,
                (month_id, cat_name, month_id, cat_name),
            )
            combined_ledger = cursor.fetchall()

            branch_data[cat_name] = {
                "base": base_amt,
                "credit": credit_amt,
                "alloc": alloc_amt,
                "spent": cat_spent,
                "ledger": combined_ledger,
            }
            grand_base += base_amt
            grand_credit += credit_amt
            grand_alloc += alloc_amt
            grand_spent += cat_spent

    txt_file = generate_txt_report(target_m, branch_data, grand_base, grand_credit, grand_alloc, grand_spent)
    pdf_file = generate_pdf_report(target_m, branch_data, grand_base, grand_credit, grand_alloc, grand_spent)

    print(f"\n{C.CYAN}═════════════════════════════════════════════════════════════════{C.RESET}")
    print(f" {C.BOLD}{C.GREEN}REPORT EXPORTED FOR: {target_m['month_name']} [{target_m['month_code']}]{C.RESET}")
    print(f"{C.CYAN}═════════════════════════════════════════════════════════════════{C.RESET}")
    print(f" {C.CYAN}•{C.RESET} Text Report : {C.YELLOW}{txt_file}{C.RESET}")
    print(f" {C.CYAN}•{C.RESET} PDF Report  : {C.YELLOW}{pdf_file}{C.RESET}")
    print(f"{C.CYAN}═════════════════════════════════════════════════════════════════{C.RESET}\n")

    try:
        os.startfile(pdf_file)
    except Exception:
        pass


# ==========================================
# 4. JUJUTSU MANUAL GENERATOR
# ==========================================
def generate_jujutsu_manual():
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    file_path = REPORTS_DIR / "yosan_command_manual.pdf"

    doc = SimpleDocTemplate(
        str(file_path),
        pagesize=A4,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36,
    )
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=20,
        leading=24,
        textColor=colors.black,
    )
    subtitle_style = ParagraphStyle(
        "DocSub",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10,
        leading=14,
        textColor=colors.black,
    )
    section_title = ParagraphStyle(
        "SecTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=12,
        leading=16,
        textColor=colors.black,
    )
    cell_head = ParagraphStyle(
        "CellHead",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=11,
        textColor=colors.white,
    )
    cmd_style = ParagraphStyle(
        "CmdStyle",
        parent=styles["Normal"],
        fontName="Courier-Bold",
        fontSize=8.5,
        leading=11,
    )
    desc_style = ParagraphStyle(
        "DescStyle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=11,
    )

    story.append(Paragraph("YOSAN CLI — MASTER COMMAND MANUAL", title_style))
    story.append(Spacer(1, 4))
    story.append(Paragraph("System Specification, Operational Flags & Continuous Entry Guide (Jujutsu Edition)", subtitle_style))
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.black, spaceAfter=14))

    commands = [
        ("yosan", "Displays the live Remaining Budget breakdown and status of the current active month."),
        ("yosan switch [flag]", "Enters continuous entry mode for a category branch. Type 'yosan -juubun' to exit."),
        ("yosan switch [flag] -d \"...\" -val X", "One-line quick entry to log an expense directly to cloud/local ledger."),
        ("yosan -u [flag] [amt] -d \"...\"", "Credits / Adds money to a branch budget and increases overall monthly allowance."),
        ("yosan -new", "Launches the Interactive Multi-Step Wizard ('p' for previous step, 'b' to cancel)."),
        ("yosan -p, yosan profile", "Opens profile dashboard (view account info, change password, delete account)."),
        ("yosan -burn", "Permanently finalizes and locks the active budget cycle into Read-Only mode."),
        ("yosan -report [code]", "Exports formatted Text and styled A4 PDF financial statements for any month."),
        ("yosan -jujutsu", "Generates and opens this Command Manual PDF documentation (Accessible logged out)."),
        ("yosan peek", "Lists all historical budget cycles stored in the active user database."),
        ("yosan peek -d [MMYYYY]", "Inspects read-only ledger and transaction history of a specific historical month."),
        ("yosan -s", "Displays total expenditure summary across all branch categories."),
        ("yosan -o", "Syncs SQLite records and launches the user-specific budget book in Excel."),
        ("yosan -logout", "Clears the active cloud token session and requires re-authentication."),
        ("yosan -delete-account", "Permanently deletes account credentials, cloud records, and local ledger files."),
    ]

    manual_data = [[Paragraph("Command / Syntax", cell_head), Paragraph("Description & Execution Behavior", cell_head)]]
    for cmd, desc in commands:
        manual_data.append([Paragraph(cmd, cmd_style), Paragraph(desc, desc_style)])

    t_manual = Table(manual_data, colWidths=[185, 340])
    t_manual.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.black),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("TOPPADDING", (0, 0), (-1, -1), 4.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ])
    )

    story.append(Paragraph("CLI REFERENCE MATRIX", section_title))
    story.append(Spacer(1, 6))
    story.append(t_manual)
    story.append(Spacer(1, 14))

    branch_data = [
        [Paragraph("Flag", cell_head), Paragraph("Category Name", cell_head), Paragraph("Default Scope", cell_head)],
        [Paragraph("-m, --mess", cmd_style), Paragraph("Mess Food", desc_style), Paragraph("Daily dining, breakfast, canteen snacks, meals", desc_style)],
        [Paragraph("-c, --clothes", cmd_style), Paragraph("Clothes", desc_style), Paragraph("Apparel, footwear, tailoring, laundry", desc_style)],
        [Paragraph("-a, --accessories", cmd_style), Paragraph("Accessories", desc_style), Paragraph("Electronics, grooming, hardware components, stationery", desc_style)],
        [Paragraph("-v, --savings", cmd_style), Paragraph("Savings", desc_style), Paragraph("Emergency reserve, long-term deposits, investment pool", desc_style)],
    ]

    t_branch = Table(branch_data, colWidths=[110, 130, 285])
    t_branch.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.black),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ])
    )
    story.append(Paragraph("BRANCH CATEGORY FLAGS", section_title))
    story.append(Spacer(1, 6))
    story.append(t_branch)
    story.append(Spacer(1, 14))

    story.append(Paragraph("INTERACTIVE WIZARD & JUUBUN LOGGING", section_title))
    story.append(Spacer(1, 4))
    story.append(
        Paragraph(
            "• <b>New Budget Wizard ('yosan -new')</b>: Supports non-destructive step navigation. Enter <b>'p'</b> or <b>'prev'</b> to return to previous steps or <b>'b'</b> / <b>'cancel'</b> to abort.<br/>"
            "• <b>Continuous Logging ('yosan switch -m')</b>: Enter sequential descriptions and amounts continuously. Type <b>'yosan -juubun'</b> or <b>'exit'</b> to finish.<br/>"
            "• <b>Real-Time Live Debouncing</b>: During login/signup, the system verifies cloud records automatically with a 1s debounce, marking valid accounts with <b>[✔]</b> without manual submission.",
            subtitle_style,
        )
    )

    doc.build(story)
    print(f"\n{C.PURPLE}═════════════════════════════════════════════════════════════════{C.RESET}")
    print(f" {C.BOLD}{C.WHITE}📖 YOSAN JUJUTSU MANUAL GENERATED SUCCESSFULLY{C.RESET}")
    print(f"{C.PURPLE}═════════════════════════════════════════════════════════════════{C.RESET}")
    print(f" {C.CYAN}•{C.RESET} PDF Manual : {C.YELLOW}{file_path}{C.RESET}")
    print(f"{C.PURPLE}═════════════════════════════════════════════════════════════════{C.RESET}\n")

    try:
        os.startfile(file_path)
    except Exception:
        pass


# ==========================================
# 5. CLI CORE FUNCTIONS
# ==========================================
def validate_month_code(code: str):
    if len(code) != 6 or not code.isdigit():
        return False, "Input must be exactly 6 digits in MMYYYY format (e.g., 082026)."
    month_num, year_num = int(code[:2]), int(code[2:])
    if month_num < 1 or month_num > 12:
        return False, f"Invalid month '{code[:2]}'. Must be between 01 and 12."
    if year_num < 2000 or year_num > 2100:
        return False, f"Invalid year '{year_num}'. Must be between 2000 and 2100."
    return True, f"{calendar.month_name[month_num]} {year_num}"


def create_new_budget():
    init_db()
    print(f"\n{C.CYAN}═══════════════════════════════════════════════════════{C.RESET}")
    print(f"        {C.BOLD}{C.WHITE}INITIALIZE NEW MONTHLY BUDGET WIZARD{C.RESET}")
    print(f"  {C.GRAY}• Type 'p' / 'prev' at any step to return to previous step{C.RESET}")
    print(f"  {C.GRAY}• Type 'b' / 'cancel' at any step to abort & exit{C.RESET}")
    print(f"{C.CYAN}═══════════════════════════════════════════════════════{C.RESET}")

    active_m = get_active_month()
    if active_m:
        print(f"{C.YELLOW}⚠️  An active budget is currently running for: {active_m['month_name']} [{active_m['month_code']}]{C.RESET}")
        burn_confirm = input(f"Are you sure you want to finalize '{active_m['month_name']}' table? (Y/N/B): ").strip().upper()

        if burn_confirm in ["B", "BACK", "CANCEL"]:
            print(f"{C.GRAY}↩ Creation canceled. Returning to terminal.{C.RESET}\n")
            return
        if burn_confirm != "Y":
            print(f"{C.RED}❌ Creation aborted. Current active table remains untouched.{C.RESET}\n")
            return

        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE months SET is_active = 0 WHERE id = ?", (active_m["id"],))
            conn.commit()

        print(f"{C.YELLOW}🔒 '{active_m['month_name']}' has been permanently archived (Read-Only).{C.RESET}")

    step = 1
    raw_code = ""
    month_display = ""
    total_income = 0.0
    branch_allocations = {}

    while True:
        # STEP 1: Month Code
        if step == 1:
            raw_input = input(f"\n{C.CYAN}[Step 1/4]{C.RESET} Enter NEW Month Code (MMYYYY, e.g. 102026) [b=cancel]: ").strip()
            if raw_input.lower() in ["b", "back", "cancel"]:
                print(f"{C.GRAY}↩ Budget creation canceled.{C.RESET}\n")
                return

            is_valid, result = validate_month_code(raw_input)
            if not is_valid:
                print(f"{C.RED}❌ {result}{C.RESET}")
                continue

            month_display = result
            raw_code = raw_input

            with get_db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT id, is_active FROM months WHERE month_code = ?", (raw_code,))
                existing_month = cursor.fetchone()

            if existing_month:
                print(f"\n{C.YELLOW}⚠️  The budget for {month_display} ({raw_code}) already exists in your database!{C.RESET}")
                if existing_month["is_active"] == 1:
                    print(f"👉 {C.GREEN}{month_display} is already active.{C.RESET}\n")
                else:
                    switch_choice = input(f"Reactivate and switch working budget to {month_display}? (Y/N): ").strip().upper()
                    if switch_choice == "Y":
                        with get_db_connection() as conn:
                            cursor = conn.cursor()
                            cursor.execute("UPDATE months SET is_active = 0")
                            cursor.execute("UPDATE months SET is_active = 1 WHERE month_code = ?", (raw_code,))
                            conn.commit()
                        sync_to_excel()
                        print(f"{C.GREEN}✅ Reactivated budget {month_display} successfully!{C.RESET}\n")
                return

            print(f" {C.GREEN}Month verified: {month_display}{C.RESET}")
            step = 2

        # STEP 2: Total Budget Amount
        elif step == 2:
            raw_input = input(f"\n{C.CYAN}[Step 2/4]{C.RESET} Enter Total Budget for {C.BOLD}{month_display}{C.RESET} (₹) [p=prev step, b=cancel]: ").strip()
            if raw_input.lower() in ["p", "prev", "previous"]:
                print(f"{C.GRAY}↩ Going back to Step 1 (Month Code)...{C.RESET}")
                step = 1
                continue
            if raw_input.lower() in ["b", "back", "cancel"]:
                print(f"{C.GRAY}↩ Budget creation canceled.{C.RESET}\n")
                return

            try:
                income_val = float(raw_input.replace(",", ""))
                if income_val <= 0:
                    print(f"{C.RED}❌ Total budget must be greater than 0.{C.RESET}")
                    continue
                total_income = income_val
                step = 3
            except ValueError:
                print(f"{C.RED}❌ Invalid input. Please enter a valid numeric amount.{C.RESET}")

        # STEP 3: Assign Branch Budgets
        elif step == 3:
            print(f"\n{C.CYAN}[Step 3/4]{C.RESET} Assign Branch Budgets for {C.BOLD}{month_display}{C.RESET} (Total: {C.GREEN}₹{total_income:,.2f}{C.RESET})")
            print(f"  {C.CYAN}[1]{C.RESET} Auto-divide by Percentage Ratio (e.g., 40%, 20%, 15%, 25%)")
            print(f"  {C.CYAN}[2]{C.RESET} Manual Entry (Enter absolute ₹ amounts per branch)")
            print(f"  {C.CYAN}[p]{C.RESET} Previous Step (Go back to change Total Budget)")
            print(f"  {C.CYAN}[b]{C.RESET} Cancel & Exit")

            choice = input(f"{C.CYAN}Select option (1, 2, p, b): {C.RESET}").strip().lower()
            if choice in ["p", "prev", "previous"]:
                print(f"{C.GRAY}↩ Going back to Step 2 (Total Budget Amount)...{C.RESET}")
                step = 2
                continue
            if choice in ["b", "back", "cancel"]:
                print(f"{C.GRAY}↩ Budget creation canceled.{C.RESET}\n")
                return

            # --- OPTION 1: PERCENTAGE ALLOCATION (With Auto-Fill on Last Branch) ---
            if choice == "1":
                print(f"\n{C.GRAY}Enter percentage for each branch (Total must = 100%, 'p'=back):{C.RESET}")
                ratios = {}
                tot_p = 0.0
                step_canceled = False
                category_items = list(CATEGORIES.values())

                for idx, cat_name in enumerate(category_items):
                    cat_color = CATEGORY_ANSI.get(cat_name, C.WHITE)
                    is_last = (idx == len(category_items) - 1)
                    remaining_p = max(0.0, round(100.0 - tot_p, 2))

                    while True:
                        if is_last and remaining_p > 0:
                            prompt_str = f"  • Percentage for [{cat_color}{cat_name:<11}{C.RESET}] (%): {remaining_p:g} (Auto-filled, Enter to accept): "
                        else:
                            prompt_str = f"  • Percentage for [{cat_color}{cat_name:<11}{C.RESET}] (%): "

                        p_in = input(prompt_str).strip()

                        if p_in.lower() in ["p", "prev", "previous", "b", "back"]:
                            step_canceled = True
                            break

                        if is_last and p_in == "":
                            p = remaining_p
                        else:
                            try:
                                p = float(p_in)
                            except ValueError:
                                print(f"    {C.RED}❌ Numeric percentage required.{C.RESET}")
                                continue

                        if p < 0:
                            print(f"    {C.RED}❌ Value cannot be negative.{C.RESET}")
                            continue

                        ratios[cat_name] = p
                        tot_p += p
                        break

                    if step_canceled:
                        break

                if step_canceled:
                    continue

                if round(tot_p, 2) == 100.0:
                    branch_allocations = {
                        cat_name: round((pct / 100.0) * total_income, 2)
                        for cat_name, pct in ratios.items()
                    }
                    step = 4
                else:
                    print(f"{C.RED}❌ Percentages sum to {tot_p:.2f}%, must equal exactly 100%.{C.RESET}")

            # --- OPTION 2: ABSOLUTE AMOUNT ALLOCATION (With Auto-Fill on Last Branch) ---
            elif choice == "2":
                print(f"\n{C.GRAY}Enter allocated ₹ amount for each branch ('p'=back):{C.RESET}")
                tot_m = 0.0
                alloc_map = {}
                step_canceled = False
                category_items = list(CATEGORIES.values())

                for idx, cat_name in enumerate(category_items):
                    cat_color = CATEGORY_ANSI.get(cat_name, C.WHITE)
                    is_last = (idx == len(category_items) - 1)
                    remaining_m = max(0.0, round(total_income - tot_m, 2))

                    while True:
                        if is_last and remaining_m > 0:
                            prompt_str = f"  • Allocated for [{cat_color}{cat_name:<11}{C.RESET}] (₹): {remaining_m:,.2f} (Auto-filled, Enter to accept): "
                        else:
                            prompt_str = f"  • Allocated for [{cat_color}{cat_name:<11}{C.RESET}] (₹): "

                        val_in = input(prompt_str).strip()

                        if val_in.lower() in ["p", "prev", "previous", "b", "back"]:
                            step_canceled = True
                            break

                        if is_last and val_in == "":
                            val = remaining_m
                        else:
                            try:
                                val = float(val_in.replace(",", ""))
                            except ValueError:
                                print(f"    {C.RED}❌ Numeric amount required.{C.RESET}")
                                continue

                        if val < 0:
                            print(f"    {C.RED}❌ Value cannot be negative.{C.RESET}")
                            continue

                        alloc_map[cat_name] = val
                        tot_m += val
                        break

                    if step_canceled:
                        break

                if step_canceled:
                    continue

                if round(tot_m, 2) == round(total_income, 2):
                    branch_allocations = alloc_map
                    step = 4
                else:
                    print(f"{C.RED}❌ Sum ₹{tot_m:,.2f} does not match total ₹{total_income:,.2f}.{C.RESET}")
            else:
                print(f"{C.RED}❌ Invalid option. Enter 1, 2, p, or b.{C.RESET}")

        # STEP 4: Review Summary & Confirmation
        elif step == 4:
            print(f"\n{C.CYAN}═══════════════════════════════════════════════════════{C.RESET}")
            print(f"  {C.BOLD}[Step 4/4] Summary to Lock In for {month_display}:{C.RESET}")
            for name, amt in branch_allocations.items():
                pct = (amt / total_income) * 100.0
                cat_color = CATEGORY_ANSI.get(name, C.WHITE)
                amt_str = f"₹{amt:,.2f}"
                print(f"    - {cat_color}{name:<12}{C.RESET}: {C.YELLOW}{amt_str:>16}{C.RESET} {C.GRAY}({pct:>5.1f}%){C.RESET}")
            total_str = f"₹{total_income:,.2f}"
            print(f"  {C.BOLD}Total Budget : {C.GREEN}{total_str:>16}{C.RESET}")
            print(f"{C.CYAN}═══════════════════════════════════════════════════════{C.RESET}")

            confirm = input(f"Confirm creation of {month_display}? (Y / p=prev / b=cancel): ").strip().upper()
            if confirm in ["P", "PREV", "PREVIOUS"]:
                print(f"{C.GRAY}↩ Going back to Step 3 (Branch Allocations)...{C.RESET}")
                step = 3
                continue
            if confirm in ["B", "BACK", "CANCEL", "N"]:
                print(f"{C.RED}❌ Budget creation canceled.{C.RESET}\n")
                return

            if confirm == "Y":
                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute("UPDATE months SET is_active = 0")
                    cursor.execute(
                        """
                        INSERT INTO months (month_code, month_name, total_budget, is_active, created_at)
                        VALUES (?, ?, ?, 1, ?)
                        """,
                        (raw_code, month_display, total_income, datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
                    )
                    month_id = cursor.lastrowid
                    for branch_name, amt in branch_allocations.items():
                        cursor.execute(
                            """
                            INSERT INTO allocations (month_id, branch, allocated_amount, base_amount)
                            VALUES (?, ?, ?, ?)
                            """,
                            (month_id, branch_name, amt, amt),
                        )
                    conn.commit()

                sync_to_excel()
                print(f"\n{C.GREEN}{C.BOLD}🎉 Budget for {month_display} is locked in and active!{C.RESET}\n")
                print_remaining_balance()
                return


def update_branch_budget(cat_name: str, add_amount: float = None, desc: str = None):
    init_db()
    active_m = get_active_month()
    if not active_m:
        print(f"\n{C.RED}🔒 ACCESS DENIED: The previous budget has been burned in and archived (Read-Only).{C.RESET}")
        print(f"👉 Run '{C.CYAN}yosan -new{C.RESET}' to create an active budget first.\n")
        return

    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT allocated_amount, base_amount FROM allocations WHERE month_id = ? AND branch = ?",
            (active_m["id"], cat_name),
        )
        row = cursor.fetchone()
        if not row:
            print(f"{C.RED}❌ Branch '{cat_name}' not found.{C.RESET}")
            return

        current_alloc = row["allocated_amount"]

        if add_amount is None:
            cat_color = CATEGORY_ANSI.get(cat_name, C.WHITE)
            print(f"\n{C.CYAN}--- Add Money / Increase Budget for [{cat_color}{cat_name}{C.RESET}{C.CYAN}] ---{C.RESET}")
            print(f"Current Allocation: {C.YELLOW}₹{current_alloc:.2f}{C.RESET}")
            while True:
                try:
                    add_amount = float(input(f"Enter amount to add to [{cat_name}] (₹): "))
                    if add_amount <= 0:
                        continue
                    break
                except ValueError:
                    print(f"{C.RED}❌ Numeric value required.{C.RESET}")

        if not desc:
            user_desc = input("Description / Reason (optional - press Enter for default): ").strip()
            desc = user_desc if user_desc else "Budget Top-up"

        new_alloc = current_alloc + add_amount
        new_total = active_m["total_budget"] + add_amount
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        cursor.execute(
            "UPDATE allocations SET allocated_amount = ? WHERE month_id = ? AND branch = ?",
            (new_alloc, active_m["id"], cat_name),
        )
        cursor.execute(
            "UPDATE months SET total_budget = ? WHERE id = ?",
            (new_total, active_m["id"]),
        )
        cursor.execute(
            "INSERT INTO topups (month_id, branch, amount, description, timestamp) VALUES (?, ?, ?, ?, ?)",
            (active_m["id"], cat_name, add_amount, desc, ts),
        )
        conn.commit()

    sync_to_excel()
    cat_color = CATEGORY_ANSI.get(cat_name, C.WHITE)
    print(f"\n{C.CYAN}═══════════════════════════════════════════════════════{C.RESET}")
    print(f" 💰 {C.BOLD}Credited Funds to [{cat_color}{cat_name}{C.RESET}{C.BOLD}] ({active_m['month_name']}){C.RESET}")
    print(f"  • Description         : {C.WHITE}{desc}{C.RESET}")
    print(f"  • Previous Allocation : {C.YELLOW}₹{current_alloc:>10.2f}{C.RESET}")
    print(f"  • Credited (+)        : {C.GREEN}+₹{add_amount:>9.2f}{C.RESET}")
    print(f"  • New Total Allocation: {C.CYAN}₹{new_alloc:>10.2f}{C.RESET}")
    print(f"  • New Total Budget    : {C.GREEN}₹{new_total:>10.2f}{C.RESET}")
    print(f"{C.CYAN}═══════════════════════════════════════════════════════{C.RESET}\n")


def add_entry(cat_name: str, desc: str, amount: float):
    init_db()
    active_m = get_active_month()
    if not active_m:
        print(f"\n{C.RED}🔒 ACCESS DENIED: Active budget has been burned in (Read-Only).{C.RESET}")
        print(f"👉 Run '{C.CYAN}yosan -new{C.RESET}' to create an active budget first.\n")
        return

    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO transactions (month_id, branch, description, amount, timestamp)
            VALUES (?, ?, ?, ?, ?)
        """,
            (active_m["id"], cat_name, desc, amount, ts),
        )
        conn.commit()

    sync_to_excel()
    cat_color = CATEGORY_ANSI.get(cat_name, C.WHITE)
    print(f" {C.GREEN}✔{C.RESET} Saved under [{cat_color}{cat_name}{C.RESET}]: {C.WHITE}{desc}{C.RESET} -> {C.RED}₹{amount:.2f}{C.RESET}")


def continuous_interactive_entry(cat_name: str):
    active_m = get_active_month()
    if not active_m:
        print(f"\n{C.RED}🔒 ACCESS DENIED: Active budget has been burned in (Read-Only).{C.RESET}")
        print(f"👉 Run '{C.CYAN}yosan -new{C.RESET}' to create an active budget first.\n")
        return

    cat_color = CATEGORY_ANSI.get(cat_name, C.WHITE)
    print(f"\n{C.CYAN}═════════════════════════════════════════════════════════════════{C.RESET}")
    print(f" 🔄 {C.BOLD}CONTINUOUS LOGGING: [{cat_color}{cat_name.upper()}{C.RESET}{C.BOLD}] ({active_m['month_name']}){C.RESET}")
    print(f"    {C.GRAY}Tip: Type 'yosan -juubun' or 'exit' to finish session.{C.RESET}")
    print(f"{C.CYAN}═════════════════════════════════════════════════════════════════{C.RESET}")

    while True:
        desc = input(f"\n[{cat_color}{cat_name}{C.RESET}] Description: ").strip()

        if desc.lower() in ["yosan -juubun", "juubun", "-juubun", "exit", "quit", "q"]:
            print(f"\n{C.GREEN}✅ Finished session for [{cat_name}]. (Juubun / 充分){C.RESET}")
            break

        if not desc:
            print(f"{C.RED}❌ Description cannot be empty.{C.RESET}")
            continue

        while True:
            amount_str = input(f"[{cat_color}{cat_name}{C.RESET}] Amount for '{desc}' (₹): ").strip()
            if amount_str.lower() in ["yosan -juubun", "juubun", "-juubun", "exit", "quit", "q"]:
                print(f"\n{C.GREEN}✅ Finished session for [{cat_name}]. (Juubun / 充分){C.RESET}")
                return
            try:
                amount = float(amount_str.replace(",", ""))
                if amount <= 0:
                    print(f"{C.RED}Amount must be greater than 0.{C.RESET}")
                    continue
                add_entry(cat_name, desc, amount)
                break
            except ValueError:
                print(f"{C.RED}❌ Invalid amount. Enter numeric ₹ value.{C.RESET}")

    print_remaining_balance()


def print_remaining_balance():
    init_db()
    active_m = get_active_month()
    username = get_current_username()

    if not active_m:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM months ORDER BY id DESC LIMIT 1")
            latest_m = cursor.fetchone()

        if not latest_m:
            print(f"\nHi {C.CYAN}{C.BOLD}{username}{C.RESET}!")
            print(f"\n{C.RED}❌ No budget records found. Run '{C.CYAN}yosan -new{C.RESET}{C.RED}' to create one.{C.RESET}")
            return

        print(f"\nHi {C.CYAN}{C.BOLD}{username}{C.RESET}!")
        print(f"\n{C.YELLOW}🔒 ALL BUDGETS ARE BURNED IN (READ-ONLY MODE){C.RESET}")
        print(f"👉 Run '{C.CYAN}yosan -new{C.RESET}' to initialize a new month's active budget.")
        peek_month_budget(latest_m["month_code"])
        return

    title_text = f"REMAINING BUDGET BREAKDOWN ({active_m['month_name'].upper()}) [ACTIVE]"
    box_width = 83

    t_spaces = max(0, box_width - len(title_text))
    t_l = t_spaces // 2
    t_r = t_spaces - t_l

    print(f"\nHi {C.CYAN}{C.BOLD}{username}{C.RESET}!")
    print(f"{C.CYAN}╔" + ("═" * box_width) + f"╗{C.RESET}")
    print(f"{C.CYAN}║{C.RESET}{' ' * t_l}{C.BOLD}{C.WHITE}{title_text}{C.RESET}{' ' * t_r}{C.CYAN}║{C.RESET}")
    print(f"{C.CYAN}╚" + ("═" * box_width) + f"╝{C.RESET}")
    print(f"{C.BOLD}{'Branch':<13} | {'Base':>11} | {'Credited':>10} | {'Total Alloc':>13} | {'Spent':>11} | {'Remaining':>12}{C.RESET}")
    print(f"{C.GRAY}───────────────────────────────────────────────────────────────────────────────────{C.RESET}")

    with get_db_connection() as conn:
        cursor = conn.cursor()
        grand_base = 0.0
        grand_credit = 0.0
        grand_alloc = 0.0
        grand_spent = 0.0

        for cat_name in CATEGORIES.values():
            cursor.execute(
                "SELECT allocated_amount, base_amount FROM allocations WHERE month_id = ? AND branch = ?",
                (active_m["id"], cat_name),
            )
            alloc_row = cursor.fetchone()
            alloc_amt = alloc_row["allocated_amount"] if alloc_row else 0.0
            base_amt = (
                alloc_row["base_amount"]
                if (alloc_row and alloc_row["base_amount"] is not None and alloc_row["base_amount"] > 0)
                else alloc_amt
            )

            cursor.execute(
                "SELECT SUM(amount) AS total FROM topups WHERE month_id = ? AND branch = ?",
                (active_m["id"], cat_name),
            )
            topup_row = cursor.fetchone()
            credit_amt = topup_row["total"] if topup_row["total"] else 0.0

            cursor.execute(
                "SELECT SUM(amount) AS total FROM transactions WHERE month_id = ? AND branch = ?",
                (active_m["id"], cat_name),
            )
            spent_row = cursor.fetchone()
            cat_spent = spent_row["total"] if spent_row["total"] else 0.0

            remaining = alloc_amt - cat_spent
            grand_base += base_amt
            grand_credit += credit_amt
            grand_alloc += alloc_amt
            grand_spent += cat_spent

            cat_color = CATEGORY_ANSI.get(cat_name, C.WHITE)
            base_str = f"₹{base_amt:>10.2f}"
            credit_str = f"+₹{credit_amt:>8.2f}" if credit_amt > 0 else f"₹{credit_amt:>8.2f}"
            alloc_str = f"₹{alloc_amt:>12.2f}"
            spent_str = f"₹{cat_spent:>10.2f}"
            rem_str = f"₹{remaining:>11.2f}" if remaining >= 0 else f"-₹{abs(remaining):>10.2f}"

            print(
                f"{cat_color}{cat_name:<13}{C.RESET} | "
                f"{C.WHITE}{base_str}{C.RESET} | "
                f"{C.GREEN if credit_amt > 0 else C.GRAY}{credit_str}{C.RESET} | "
                f"{C.CYAN}{alloc_str}{C.RESET} | "
                f"{C.RED}{spent_str}{C.RESET} | "
                f"{C.GREEN if remaining >= 0 else C.RED}{rem_str}{C.RESET}"
            )

        print(f"{C.GRAY}───────────────────────────────────────────────────────────────────────────────────{C.RESET}")
        total_rem = grand_alloc - grand_spent
        tb_str = f"₹{grand_base:>10.2f}"
        tc_str = f"+₹{grand_credit:>8.2f}" if grand_credit > 0 else f"₹{grand_credit:>8.2f}"
        ta_str = f"₹{grand_alloc:>12.2f}"
        ts_str = f"₹{grand_spent:>10.2f}"
        tr_str = f"₹{total_rem:>11.2f}" if total_rem >= 0 else f"-₹{abs(total_rem):>10.2f}"

        print(
            f"{C.BOLD}{'TOTAL':<13}{C.RESET} | "
            f"{C.YELLOW}{tb_str}{C.RESET} | "
            f"{C.GREEN if grand_credit > 0 else C.GRAY}{tc_str}{C.RESET} | "
            f"{C.CYAN}{ta_str}{C.RESET} | "
            f"{C.RED}{ts_str}{C.RESET} | "
            f"{C.GREEN if total_rem >= 0 else C.RED}{C.BOLD}{tr_str}{C.RESET}"
        )
        print(f"{C.CYAN}═══════════════════════════════════════════════════════════════════════════════════{C.RESET}\n")


def print_summary():
    init_db()
    active_m = get_active_month()
    if not active_m:
        print(f"\n{C.RED}❌ No active budget found. Run '{C.CYAN}yosan -new{C.RESET}{C.RED}' first.{C.RESET}")
        return

    print(f"\n{C.CYAN}═════════════════════════════════════════════{C.RESET}")
    print(f"       {C.BOLD}{C.WHITE}YOSAN SPENDING SUMMARY ({active_m['month_name']}){C.RESET}")
    print(f"{C.CYAN}═════════════════════════════════════════════{C.RESET}")

    with get_db_connection() as conn:
        cursor = conn.cursor()
        grand_total = 0.0
        for cat_name in CATEGORIES.values():
            cursor.execute(
                "SELECT SUM(amount) AS total FROM transactions WHERE month_id = ? AND branch = ?",
                (active_m["id"], cat_name),
            )
            row = cursor.fetchone()
            cat_total = row["total"] if row["total"] else 0.0
            cat_color = CATEGORY_ANSI.get(cat_name, C.WHITE)
            print(f" * {cat_color}{cat_name:<20}{C.RESET}: {C.RED}₹{cat_total:>10.2f}{C.RESET}")
            grand_total += cat_total

        print(f"{C.GRAY}─────────────────────────────────────────────{C.RESET}")
        print(f" {C.BOLD}GRAND TOTAL EXPENDITURE:{C.RESET} {C.RED}{C.BOLD}₹{grand_total:>10.2f}{C.RESET}")
        print(f"{C.CYAN}═════════════════════════════════════════════{C.RESET}\n")


def peek_month_budget(month_code: str):
    init_db()
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM months WHERE month_code = ?", (month_code,))
        target_m = cursor.fetchone()

    if not target_m:
        print(f"\n{C.RED}❌ No budget record found for Month Code: '{month_code}'.{C.RESET}")
        print(f"{C.GRAY}Tip: Run 'yosan peek' without flags to view all available months.{C.RESET}")
        return

    active_tag = f" {C.GREEN}(CURRENT ACTIVE){C.RESET}" if target_m["is_active"] == 1 else f" {C.YELLOW}(ARCHIVED - READ ONLY){C.RESET}"
    title_text = f"PEEK HISTORICAL BUDGET: {target_m['month_name']} [{month_code}]"
    box_width = 83

    t_spaces = max(0, box_width - len(title_text))
    t_l = t_spaces // 2
    t_r = t_spaces - t_l

    print(f"\n{C.CYAN}╔" + ("═" * box_width) + f"╗{C.RESET}")
    print(f"{C.CYAN}║{C.RESET}{' ' * t_l}{C.BOLD}{C.WHITE}{title_text}{C.RESET}{' ' * t_r}{C.CYAN}║{C.RESET}")
    print(f"{C.CYAN}╚" + ("═" * box_width) + f"╝{C.RESET}")
    print(f"Status: {active_tag}")
    print(f"{C.BOLD}{'Branch':<13} | {'Base':>11} | {'Credited':>10} | {'Total Alloc':>13} | {'Spent':>11} | {'Remaining':>12}{C.RESET}")
    print(f"{C.GRAY}───────────────────────────────────────────────────────────────────────────────────{C.RESET}")

    with get_db_connection() as conn:
        cursor = conn.cursor()
        grand_base = 0.0
        grand_credit = 0.0
        grand_alloc = 0.0
        grand_spent = 0.0

        for cat_name in CATEGORIES.values():
            cursor.execute(
                "SELECT allocated_amount, base_amount FROM allocations WHERE month_id = ? AND branch = ?",
                (target_m["id"], cat_name),
            )
            alloc_row = cursor.fetchone()
            alloc_amt = alloc_row["allocated_amount"] if alloc_row else 0.0
            base_amt = (
                alloc_row["base_amount"]
                if (alloc_row and alloc_row["base_amount"] is not None and alloc_row["base_amount"] > 0)
                else alloc_amt
            )

            cursor.execute(
                "SELECT SUM(amount) AS total FROM topups WHERE month_id = ? AND branch = ?",
                (target_m["id"], cat_name),
            )
            topup_row = cursor.fetchone()
            credit_amt = topup_row["total"] if topup_row["total"] else 0.0

            cursor.execute(
                "SELECT SUM(amount) AS total FROM transactions WHERE month_id = ? AND branch = ?",
                (target_m["id"], cat_name),
            )
            stats_row = cursor.fetchone()
            cat_spent = stats_row["total"] if stats_row["total"] else 0.0

            remaining = alloc_amt - cat_spent
            grand_base += base_amt
            grand_credit += credit_amt
            grand_alloc += alloc_amt
            grand_spent += cat_spent

            cat_color = CATEGORY_ANSI.get(cat_name, C.WHITE)
            base_str = f"₹{base_amt:>10.2f}"
            credit_str = f"+₹{credit_amt:>8.2f}" if credit_amt > 0 else f"₹{credit_amt:>8.2f}"
            alloc_str = f"₹{alloc_amt:>12.2f}"
            spent_str = f"₹{cat_spent:>10.2f}"
            rem_str = f"₹{remaining:>11.2f}" if remaining >= 0 else f"-₹{abs(remaining):>10.2f}"

            print(
                f"{cat_color}{cat_name:<13}{C.RESET} | "
                f"{C.WHITE}{base_str}{C.RESET} | "
                f"{C.GREEN if credit_amt > 0 else C.GRAY}{credit_str}{C.RESET} | "
                f"{C.CYAN}{alloc_str}{C.RESET} | "
                f"{C.RED}{spent_str}{C.RESET} | "
                f"{C.GREEN if remaining >= 0 else C.RED}{rem_str}{C.RESET}"
            )

        print(f"{C.GRAY}───────────────────────────────────────────────────────────────────────────────────{C.RESET}")
        total_rem = grand_alloc - grand_spent
        tb_str = f"₹{grand_base:>10.2f}"
        tc_str = f"+₹{grand_credit:>8.2f}" if grand_credit > 0 else f"₹{grand_credit:>8.2f}"
        ta_str = f"₹{grand_alloc:>12.2f}"
        ts_str = f"₹{grand_spent:>10.2f}"
        tr_str = f"₹{total_rem:>11.2f}" if total_rem >= 0 else f"-₹{abs(total_rem):>10.2f}"

        print(
            f"{C.BOLD}{'TOTAL':<13}{C.RESET} | "
            f"{C.YELLOW}{tb_str}{C.RESET} | "
            f"{C.GREEN if grand_credit > 0 else C.GRAY}{tc_str}{C.RESET} | "
            f"{C.CYAN}{ta_str}{C.RESET} | "
            f"{C.RED}{ts_str}{C.RESET} | "
            f"{C.GREEN if total_rem >= 0 else C.RED}{C.BOLD}{tr_str}{C.RESET}"
        )
        print(f"{C.CYAN}═══════════════════════════════════════════════════════════════════════════════════{C.RESET}")

        cursor.execute(
            """
            SELECT timestamp, 'Expense' AS type, branch, description, amount FROM transactions WHERE month_id = ?
            UNION ALL
            SELECT timestamp, 'Credit' AS type, branch, description, amount FROM topups WHERE month_id = ?
            ORDER BY timestamp ASC
            """,
            (target_m["id"], target_m["id"]),
        )
        tx_list = cursor.fetchall()
        if tx_list:
            print(f"\n{C.CYAN}--- Itemized Transaction Log ({len(tx_list)} entries) ---{C.RESET}")
            for idx, tx in enumerate(tx_list, start=1):
                amt_str = f"{C.GREEN}+₹{tx['amount']:.2f}{C.RESET}" if tx["type"] == "Credit" else f"{C.RED}₹{tx['amount']:.2f}{C.RESET}"
                cat_color = CATEGORY_ANSI.get(tx["branch"], C.WHITE)
                print(f" {C.GRAY}{idx:>2}.{C.RESET} [{C.GRAY}{tx['timestamp']}{C.RESET}] [{C.BOLD}{tx['type']:<7}{C.RESET}] [{cat_color}{tx['branch']:<11}{C.RESET}] {tx['description']} -> {amt_str}")
            print(f"{C.GRAY}───────────────────────────────────────────────────────────────────────────────────{C.RESET}\n")
        else:
            print(f"\n{C.GRAY}ℹ️  No transactions recorded for this month.{C.RESET}\n")


def list_all_available_months():
    init_db()
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT month_code, month_name, total_budget, is_active FROM months ORDER BY id DESC")
        rows = cursor.fetchall()

    if not rows:
        print(f"\n{C.RED}❌ No months found in database.{C.RESET}")
        return

    print(f"\n{C.CYAN}═══════════════════════════════════════════════════════{C.RESET}")
    print(f"              {C.BOLD}{C.WHITE}AVAILABLE STORED MONTHS{C.RESET}")
    print(f"{C.CYAN}═══════════════════════════════════════════════════════{C.RESET}")
    print(f"{C.BOLD}{'Code':<8} | {'Month Name':<18} | {'Budget':>11} | {'Status'}{C.RESET}")
    print(f"{C.GRAY}───────────────────────────────────────────────────────{C.RESET}")
    for r in rows:
        status = f"{C.GREEN}Active{C.RESET}" if r["is_active"] == 1 else f"{C.GRAY}Archived{C.RESET}"
        print(f"{C.YELLOW}{r['month_code']:<8}{C.RESET} | {r['month_name']:<18} | {C.GREEN}₹{r['total_budget']:>10.2f}{C.RESET} | {status}")
    print(f"{C.CYAN}═══════════════════════════════════════════════════════{C.RESET}")
    print(f"{C.GRAY}Run: yosan peek -d <code (e.g. 082026)>{C.RESET}\n")


def burn_current_budget():
    init_db()
    active_m = get_active_month()
    if not active_m:
        print(f"\n{C.YELLOW}🔒 There is no active budget running. Everything is already burned in / read-only.{C.RESET}")
        return

    print(f"\n{C.YELLOW}⚠️  Current Active Budget: {active_m['month_name']} [{active_m['month_code']}]{C.RESET}")
    confirm = input(f"Are you sure you want to {C.RED}BURN IN{C.RESET} and lock '{active_m['month_name']}' permanently? (Y/N): ").strip().upper()
    if confirm == "Y":
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE months SET is_active = 0 WHERE id = ?", (active_m["id"],))
            conn.commit()
        print(f"{C.GREEN} '{active_m['month_name']}' is now BURNED IN and set to READ-ONLY.{C.RESET}\n")
    else:
        print(f"{C.GRAY}❌ Action canceled. Active budget remains untouched.{C.RESET}\n")


# ==========================================
# 6. MAIN ROUTING
# ==========================================
def main():
    # 1. Public Command: Jujutsu Manual (Available even if logged out)
    if "-jujutsu" in sys.argv or "--jujutsu" in sys.argv or "jujutsu" in sys.argv:
        generate_jujutsu_manual()
        return

    # 2. Handle manual logout
    if "-logout" in sys.argv or "--logout" in sys.argv:
        clear_session()
        return

    # 3. Handle manual account deletion
    if "-delete-account" in sys.argv or "--delete-account" in sys.argv or "-delete" in sys.argv:
        delete_account_flow()
        return

    # 4. Master Authentication Gate
    if not require_login():
        sys.exit(1)

    # 5. Intercept profile shortcut
    if "-p" in sys.argv or "--profile" in sys.argv or "profile" in sys.argv:
        show_profile_view()
        return

    init_db()

    # 6. Intercept report command: yosan -report [optional_month_code]
    if "-report" in sys.argv or "--report" in sys.argv:
        target_code = None
        for arg in sys.argv[1:]:
            if arg not in ["-report", "--report"]:
                target_code = arg
                break
        export_monthly_report(target_code)
        return

    # 7. Intercept quick update / top-up: yosan -u [branch] [amount] -d [description]
    if "-u" in sys.argv or "--update" in sys.argv:
        active_m = get_active_month()
        if not active_m:
            print(f"\n{C.RED}🔒 ACCESS DENIED: Active budget has been burned in (Read-Only).{C.RESET}")
            print(f"👉 Run '{C.CYAN}yosan -new{C.RESET}' to create an active budget first.\n")
            return

        target_cat = None
        amount_val = None
        desc_val = None

        if "-d" in sys.argv:
            d_idx = sys.argv.index("-d")
            if d_idx + 1 < len(sys.argv):
                desc_val = sys.argv[d_idx + 1]

        if "-m" in sys.argv:
            target_cat = "Mess Food"
        elif "-c" in sys.argv:
            target_cat = "Clothes"
        elif "-a" in sys.argv:
            target_cat = "Accessories"
        elif "-v" in sys.argv:
            target_cat = "Savings"

        for arg in sys.argv[1:]:
            if arg not in [
                "-u", "--update", "-m", "--mess", "-c", "--clothes",
                "-a", "--accessories", "-v", "--savings", "-d", desc_val,
            ]:
                try:
                    amount_val = float(arg.replace(",", ""))
                    break
                except ValueError:
                    pass

        if target_cat:
            update_branch_budget(target_cat, amount_val, desc_val)
        else:
            print(f"\n{C.CYAN}Select a branch to add money to:{C.RESET}")
            print(f"  {C.CYAN}[1]{C.RESET} Mess Food (-m)")
            print(f"  {C.CYAN}[2]{C.RESET} Clothes (-c)")
            print(f"  {C.CYAN}[3]{C.RESET} Accessories (-a)")
            print(f"  {C.CYAN}[4]{C.RESET} Savings (-v)")
            ch = input(f"{C.CYAN}Enter choice (1-4): {C.RESET}").strip()
            mapping = {
                "1": "Mess Food",
                "2": "Clothes",
                "3": "Accessories",
                "4": "Savings",
            }
            if ch in mapping:
                update_branch_budget(mapping[ch], amount_val, desc_val)
            else:
                print(f"{C.RED}❌ Invalid selection.{C.RESET}")
        return

    parser = argparse.ArgumentParser(description="Yosan Budget CLI (DB Engine)")
    parser.add_argument("-new", "--new", action="store_true", help="Initialize a new monthly budget")
    parser.add_argument("-burn", "--burn", action="store_true", help="Burn in / finalize the active budget to read-only")
    parser.add_argument("-jujutsu", "--jujutsu", action="store_true", help="Generate Command Reference Manual PDF")
    parser.add_argument("-s", "--summary", action="store_true", help="View sum total balance spent")
    parser.add_argument("-sh", "--show-remaining", action="store_true", help="View remaining branch allowance")
    parser.add_argument("-o", "--open", action="store_true", help="Open budget book in Excel")
    parser.add_argument("-p", "--profile", action="store_true", help="View profile and account settings")
    parser.add_argument("-logout", "--logout", action="store_true", help="Log out active session")
    parser.add_argument("-delete-account", "--delete-account", action="store_true", help="Permanently delete active account and data")

    subparsers = parser.add_subparsers(dest="subcommand")

    subparsers.add_parser("profile", help="Manage user profile and security settings")
    subparsers.add_parser("jujutsu", help="Open the Command Manual PDF")

    switch_parser = subparsers.add_parser("switch", help="Switch to a specific category")
    switch_parser.add_argument("-m", "--mess", action="store_true", help="Mess Food")
    switch_parser.add_argument("-c", "--clothes", action="store_true", help="Clothes")
    switch_parser.add_argument("-a", "--accessories", action="store_true", help="Accessories")
    switch_parser.add_argument("-v", "--savings", action="store_true", help="Savings")
    switch_parser.add_argument("-d", "--desc", help="Description")
    switch_parser.add_argument("-val", "--amount", type=float, help="Amount")

    peek_parser = subparsers.add_parser("peek", help="Inspect read-only historical budget data")
    peek_parser.add_argument("-d", "--date", help="Month code to peek (e.g., 082026)")

    args = parser.parse_args()

    if args.logout:
        clear_session()
        return

    if args.delete_account:
        delete_account_flow()
        return

    if args.profile or args.subcommand == "profile":
        show_profile_view()
        return

    if args.jujutsu or args.subcommand == "jujutsu":
        generate_jujutsu_manual()
        return

    if args.burn:
        burn_current_budget()
        return

    if args.new:
        create_new_budget()
        return

    if args.show_remaining:
        print_remaining_balance()
        return

    if args.summary:
        print_summary()
        return

    if args.open:
        sync_to_excel()
        try:
            os.startfile(get_user_excel_path())
        except Exception:
            pass
        return

    if args.subcommand == "peek":
        if args.date:
            peek_month_budget(args.date)
        else:
            list_all_available_months()
        return

    if args.subcommand == "switch":
        active_m = get_active_month()
        if not active_m:
            print(f"\n{C.RED}🔒 ACCESS DENIED: Active budget has been burned in (Read-Only).{C.RESET}")
            print(f"👉 Run '{C.CYAN}yosan -new{C.RESET}' to create an active budget first.\n")
            return

        target_cat = None
        if args.mess:
            target_cat = "Mess Food"
        elif args.clothes:
            target_cat = "Clothes"
        elif args.accessories:
            target_cat = "Accessories"
        elif args.savings:
            target_cat = "Savings"

        if target_cat:
            if args.desc and args.amount is not None:
                add_entry(target_cat, args.desc, args.amount)
            else:
                continuous_interactive_entry(target_cat)
        else:
            print(f"{C.RED}Please specify a category flag: -m, -c, -a, or -v{C.RESET}")
        return

    print_remaining_balance()


if __name__ == "__main__":
    main()
