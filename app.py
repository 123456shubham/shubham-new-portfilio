from __future__ import annotations

from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo
import base64
import json
import html as html_lib
import os
import re
import smtplib
import sqlite3
import threading

import gspread
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from dotenv import load_dotenv
from flask import (
    Flask,
    Response,
    flash,
    redirect,
    render_template,
    request,
    stream_with_context,
    url_for,
)

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None  # type: ignore[assignment]


BASE_DIR = Path(__file__).resolve().parent
DATABASE_PATH = BASE_DIR / "portfolio.db"
load_dotenv(BASE_DIR / ".env")

app = Flask(__name__)
app.config.update(
    SECRET_KEY=os.getenv("SECRET_KEY", "development-only-change-me"),
    MAX_CONTENT_LENGTH=12 * 1024 * 1024,
)

MAIL_USERNAME = os.getenv("MAIL_USERNAME", "").strip()
MAIL_PASSWORD = os.getenv("MAIL_PASSWORD", "").replace(" ", "").strip()
MAIL_RECEIVER = os.getenv(
    "MAIL_RECEIVER", "princechauhan31081997@gmail.com"
).strip()
ENQUIRY_EXCEL_PATH = BASE_DIR / os.getenv(
    "ENQUIRY_EXCEL_FILE", "Portfolio_Enquiries.xlsx"
).strip()
EXCEL_LOCK = threading.Lock()
LOGO_PATH = BASE_DIR / "static" / "images" / "shubham-chauhan-logo.png"
INDIA_TIMEZONE = ZoneInfo("Asia/Kolkata")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "").strip()
GOOGLE_SHEET_NAME = os.getenv("GOOGLE_SHEET_NAME", "Enquiries").strip()
GOOGLE_SERVICE_ACCOUNT_PATH = BASE_DIR / os.getenv(
    "GOOGLE_SERVICE_ACCOUNT_FILE", "google-service-account.json"
).strip()
GOOGLE_SHEET_LOCK = threading.Lock()

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-5-mini").strip()

WHATSAPP_NUMBER = re.sub(
    r"\D",
    "",
    os.getenv("WHATSAPP_NUMBER", "918929932706"),
)

openai_client = (
    OpenAI(api_key=OPENAI_API_KEY)
    if OpenAI is not None and OPENAI_API_KEY
    else None
)


def database_connection() -> sqlite3.Connection:
    connection = sqlite3.connect(DATABASE_PATH)
    connection.row_factory = sqlite3.Row
    return connection


def initialise_database() -> None:
    with database_connection() as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS contact_messages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                email TEXT NOT NULL,
                phone TEXT,
                subject TEXT NOT NULL,
                message TEXT NOT NULL,
                email_status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT NOT NULL
            )
            """
        )
        connection.commit()


def safe_text(value: Any, maximum: int) -> str:
    return str(value or "").strip()[:maximum]


def looks_like_email(value: str) -> bool:
    return bool(
        re.fullmatch(
            r"[^@\s]+@[^@\s]+\.[^@\s]+",
            value,
        )
    )


def format_india_datetime(value: str | datetime | None = None) -> str:
    """Return a consistent, human-readable timestamp in Indian Standard Time."""
    if value is None:
        moment = datetime.now(timezone.utc)
    elif isinstance(value, datetime):
        moment = value
    else:
        moment = datetime.fromisoformat(value.replace("Z", "+00:00"))
    if moment.tzinfo is None:
        moment = moment.replace(tzinfo=timezone.utc)
    return moment.astimezone(INDIA_TIMEZONE).strftime("%d %b %Y, %I:%M %p IST")


EXCEL_HEADERS = [
    "Enquiry ID",
    "Received At",
    "Client Name",
    "Email Address",
    "Phone / WhatsApp",
    "Project Subject",
    "Project Details",
    "Lead Status",
    "Email Delivery",
    "Source",
    "Enquiry Validity",
    "Validation Notes",
]


def assess_enquiry_validity(email: str, phone: str, message: str) -> tuple[str, str]:
    """Classify basic lead quality without rejecting legitimate optional-phone leads."""
    notes: list[str] = []
    if not looks_like_email(email):
        return "Invalid", "Invalid email address"

    phone_digits = re.sub(r"\D", "", phone)
    if phone and not 7 <= len(phone_digits) <= 15:
        notes.append("Phone number needs review")
    if len(message.strip()) < 20:
        notes.append("Project details are very short")
    if len(re.findall(r"https?://|www\.", message, flags=re.IGNORECASE)) >= 3:
        notes.append("Multiple links detected")

    if any(note in notes for note in ("Phone number needs review", "Multiple links detected")):
        return "Review", "; ".join(notes)
    return "Valid", "; ".join(notes) or "Email and enquiry details passed validation"


def style_enquiry_sheet(sheet) -> None:
    """Apply a polished, reusable layout to the enquiry workbook."""
    navy = "172554"
    blue = "2563EB"
    white = "FFFFFF"
    muted = "64748B"
    border_colour = "CBD5E1"

    for merged_range in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(merged_range))
    sheet._images = []
    sheet.conditional_formatting._cf_rules.clear()
    sheet.data_validations.dataValidation = []

    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A5"
    sheet.merge_cells("A1:B3")
    sheet.merge_cells("C1:L1")
    sheet["C1"] = "PORTFOLIO ENQUIRIES"
    sheet["C1"].font = Font(name="Aptos Display", size=20, bold=True, color=white)
    sheet["C1"].alignment = Alignment(horizontal="left", vertical="center")

    sheet.merge_cells("C2:L2")
    sheet["C2"] = "Professional lead register | Shubham Chauhan, Software Engineer"
    sheet["C2"].font = Font(name="Aptos", size=10, italic=True, color="CBD5E1")
    sheet["C2"].alignment = Alignment(vertical="center")
    for row in sheet.iter_rows(min_row=1, max_row=3, min_col=1, max_col=12):
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=navy)
    sheet.row_dimensions[1].height = 38
    sheet.row_dimensions[2].height = 24
    sheet.row_dimensions[3].height = 8

    if LOGO_PATH.exists():
        logo = ExcelImage(LOGO_PATH)
        logo.width = 88
        logo.height = 88
        sheet.add_image(logo, "A1")

    thin_border = Border(bottom=Side(style="thin", color=border_colour))
    for column, header in enumerate(EXCEL_HEADERS, 1):
        cell = sheet.cell(row=4, column=column, value=header)
        cell.font = Font(name="Aptos", size=10, bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=blue)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = thin_border
    sheet.row_dimensions[4].height = 28
    sheet.auto_filter.ref = "A4:L4"

    widths = {
        "A": 13, "B": 22, "C": 22, "D": 30, "E": 20,
        "F": 30, "G": 55, "H": 18, "I": 18, "J": 18,
        "K": 18, "L": 38,
    }
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width

    status_validation = DataValidation(
        type="list",
        formula1='"New,Contacted,Qualified,Proposal Sent,Won,Lost,Closed"',
        allow_blank=False,
    )
    status_validation.promptTitle = "Lead status"
    status_validation.prompt = "Select the current enquiry stage."
    status_validation.error = "Please select a status from the list."
    status_validation.errorTitle = "Invalid status"
    status_validation.showErrorMessage = True
    status_validation.showInputMessage = True
    sheet.add_data_validation(status_validation)
    status_validation.add("H5:H1048576")

    validity_colours = {"Valid": "DCFCE7", "Review": "FEF3C7", "Invalid": "FEE2E2"}
    for validity, colour in validity_colours.items():
        sheet.conditional_formatting.add(
            "K5:K1048576",
            FormulaRule(
                formula=[f'$K5="{validity}"'],
                fill=PatternFill("solid", fgColor=colour),
                font=Font(name="Aptos", bold=True, color=navy),
            ),
        )

    row_validity_colours = {"Valid": "F0FDF4", "Review": "FFFBEB", "Invalid": "FEF2F2"}
    for validity, colour in row_validity_colours.items():
        sheet.conditional_formatting.add(
            "A5:L1048576",
            FormulaRule(
                formula=[f'$K5="{validity}"'],
                fill=PatternFill("solid", fgColor=colour),
            ),
        )

    status_colours = {
        "New": "DBEAFE",
        "Contacted": "FEF3C7",
        "Qualified": "EDE9FE",
        "Proposal Sent": "FFEDD5",
        "Won": "DCFCE7",
        "Lost": "FEE2E2",
        "Closed": "E2E8F0",
    }
    for status, colour in status_colours.items():
        sheet.conditional_formatting.add(
            "H5:H1048576",
            FormulaRule(
                formula=[f'$H5="{status}"'],
                fill=PatternFill("solid", fgColor=colour),
                font=Font(name="Aptos", bold=True, color=navy),
            ),
        )


def append_enquiry_to_excel(
    *,
    enquiry_id: int,
    created_at: str,
    name: str,
    email: str,
    phone: str,
    subject: str,
    message: str,
    email_status: str,
) -> None:
    """Create the formatted workbook when needed and append one enquiry safely."""
    with EXCEL_LOCK:
        if ENQUIRY_EXCEL_PATH.exists():
            workbook = load_workbook(ENQUIRY_EXCEL_PATH)
            sheet = workbook["Enquiries"] if "Enquiries" in workbook.sheetnames else workbook.create_sheet("Enquiries")
            current_headers = [sheet.cell(4, column).value for column in range(1, len(EXCEL_HEADERS) + 1)]
            if current_headers != EXCEL_HEADERS:
                style_enquiry_sheet(sheet)
        else:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Enquiries"
            style_enquiry_sheet(sheet)

        row = max(sheet.max_row + 1, 5)
        validity, validation_notes = assess_enquiry_validity(email, phone, message)
        values = [
            f"ENQ-{enquiry_id:05d}",
            format_india_datetime(created_at),
            name,
            email,
            phone or "Not provided",
            subject,
            message,
            "New",
            "Sent" if email_status == "sent" else "Saved only",
            "Portfolio Website",
            validity,
            validation_notes,
        ]
        for column, value in enumerate(values, 1):
            cell = sheet.cell(row=row, column=column, value=value)
            cell.font = Font(name="Aptos", size=10, color="1E293B")
            cell.alignment = Alignment(vertical="top", wrap_text=column in {6, 7, 12})
            cell.border = Border(bottom=Side(style="hair", color="E2E8F0"))
            if row % 2 == 0 and column != 8:
                cell.fill = PatternFill("solid", fgColor="F8FAFC")
        sheet.row_dimensions[row].height = 42
        sheet.auto_filter.ref = f"A4:L{row}"
        style_excel_lead_finder(workbook)
        workbook.save(ENQUIRY_EXCEL_PATH)


def style_excel_lead_finder(workbook) -> None:
    """Create a premium Enquiry ID lookup sheet for desktop Excel users."""
    if "Lead Finder" in workbook.sheetnames:
        del workbook["Lead Finder"]
    sheet = workbook.create_sheet("Lead Finder", 0)
    sheet.sheet_view.showGridLines = False
    sheet.merge_cells("A1:L1")
    sheet["A1"] = "LEAD FINDER"
    sheet["A1"].font = Font(name="Aptos Display", size=20, bold=True, color="FFFFFF")
    sheet["A1"].fill = PatternFill("solid", fgColor="172554")
    sheet["A1"].alignment = Alignment(vertical="center")
    sheet.row_dimensions[1].height = 38
    sheet["A3"] = "Search by Enquiry ID"
    sheet["A3"].font = Font(name="Aptos", bold=True, color="172554")
    sheet["B3"] = ""
    sheet["B3"].fill = PatternFill("solid", fgColor="EFF6FF")
    sheet["B3"].border = Border(bottom=Side(style="medium", color="2563EB"))
    validation = DataValidation(type="list", formula1="'Enquiries'!$A$5:$A$1048576", allow_blank=True)
    sheet.add_data_validation(validation)
    validation.add(sheet["B3"])
    for column, header in enumerate(EXCEL_HEADERS, 1):
        cell = sheet.cell(5, column, header)
        cell.font = Font(name="Aptos", bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2563EB")
    sheet["A6"] = '=IF(B3="","Select an Enquiry ID above",IFERROR(XLOOKUP(B3,Enquiries!A:A,Enquiries!A:L),"No enquiry found"))'
    sheet["A6"].alignment = Alignment(wrap_text=True, vertical="top")
    for column, width in {"A":16,"B":23,"C":24,"D":30,"E":20,"F":30,"G":50,"H":18,"I":18,"J":20,"K":18,"L":38}.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A5"


def append_enquiry_to_google_sheet(
    *,
    enquiry_id: int,
    created_at: str,
    name: str,
    email: str,
    phone: str,
    subject: str,
    message: str,
    email_status: str,
) -> bool:
    """Append an enquiry to Google Sheets immediately when credentials exist."""
    if not GOOGLE_SHEET_ID or not GOOGLE_SERVICE_ACCOUNT_PATH.is_file():
        app.logger.warning("Google Sheets credentials are missing; live sync skipped.")
        return False

    with GOOGLE_SHEET_LOCK:
        client = gspread.service_account(filename=str(GOOGLE_SERVICE_ACCOUNT_PATH))
        spreadsheet = client.open_by_key(GOOGLE_SHEET_ID)
        try:
            sheet = spreadsheet.worksheet(GOOGLE_SHEET_NAME)
        except gspread.WorksheetNotFound:
            sheet = spreadsheet.add_worksheet(
                title=GOOGLE_SHEET_NAME,
                rows=1000,
                cols=len(EXCEL_HEADERS),
            )

        if sheet.row_values(1) != EXCEL_HEADERS:
            sheet.update(range_name="A1:L1", values=[EXCEL_HEADERS])
            sheet.freeze(rows=1)
            sheet.format(
                "A1:L1",
                {
                    "backgroundColor": {"red": 0.145, "green": 0.388, "blue": 0.922},
                    "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                    "horizontalAlignment": "LEFT",
                },
            )
            sheet.set_basic_filter("A1:L1000")
            spreadsheet.batch_update(
                {
                    "requests": [
                        {
                            "updateDimensionProperties": {
                                "range": {"sheetId": sheet.id, "dimension": "ROWS", "startIndex": 0, "endIndex": 1},
                                "properties": {"pixelSize": 42},
                                "fields": "pixelSize",
                            }
                        },
                        {
                            "updateDimensionProperties": {
                                "range": {"sheetId": sheet.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 12},
                                "properties": {"pixelSize": 170},
                                "fields": "pixelSize",
                            }
                        },
                        {
                            "updateDimensionProperties": {
                                "range": {"sheetId": sheet.id, "dimension": "COLUMNS", "startIndex": 6, "endIndex": 7},
                                "properties": {"pixelSize": 360},
                                "fields": "pixelSize",
                            }
                        },
                        {
                            "repeatCell": {
                                "range": {"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 12},
                                "cell": {"userEnteredFormat": {"verticalAlignment": "TOP", "wrapStrategy": "WRAP", "padding": {"top": 8, "bottom": 8, "left": 8, "right": 8}}},
                                "fields": "userEnteredFormat(verticalAlignment,wrapStrategy,padding)",
                            }
                        },
                        {
                            "setDataValidation": {
                                "range": {"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": 7, "endColumnIndex": 8},
                                "rule": {
                                    "condition": {"type": "ONE_OF_LIST", "values": [{"userEnteredValue": value} for value in ["New", "Contacted", "Qualified", "Proposal Sent", "Won", "Lost", "Closed"]]},
                                    "strict": True,
                                    "showCustomUi": True,
                                },
                            }
                        },
                        *[
                            {
                                "addConditionalFormatRule": {
                                    "rule": {
                                        "ranges": [{"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": 7, "endColumnIndex": 8}],
                                        "booleanRule": {
                                            "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": status}]},
                                            "format": {"backgroundColor": colour, "textFormat": {"bold": True}},
                                        },
                                    },
                                    "index": 0,
                                }
                            }
                            for status, colour in {
                                "New": {"red": .86, "green": .92, "blue": .99},
                                "Contacted": {"red": .99, "green": .94, "blue": .76},
                                "Qualified": {"red": .92, "green": .89, "blue": .99},
                                "Proposal Sent": {"red": 1, "green": .91, "blue": .82},
                                "Won": {"red": .86, "green": .97, "blue": .89},
                                "Lost": {"red": .99, "green": .88, "blue": .88},
                                "Closed": {"red": .89, "green": .91, "blue": .94},
                            }.items()
                        ],
                        *[
                            {
                                "addConditionalFormatRule": {
                                    "rule": {
                                        "ranges": [{"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": 10, "endColumnIndex": 11}],
                                        "booleanRule": {
                                            "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": validity}]},
                                            "format": {"backgroundColor": colour, "textFormat": {"bold": True}},
                                        },
                                    },
                                    "index": 0,
                                }
                            }
                            for validity, colour in {
                                "Valid": {"red": .86, "green": .97, "blue": .89},
                                "Review": {"red": .99, "green": .94, "blue": .76},
                                "Invalid": {"red": .99, "green": .88, "blue": .88},
                            }.items()
                        ],
                        *[
                            {
                                "addConditionalFormatRule": {
                                    "rule": {
                                        "ranges": [{"sheetId": sheet.id, "startRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 12}],
                                        "booleanRule": {
                                            "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": f'=$K2="{validity}"'}]},
                                            "format": {"backgroundColor": colour},
                                        },
                                    },
                                    "index": 0,
                                }
                            }
                            for validity, colour in {
                                "Valid": {"red": .94, "green": .99, "blue": .96},
                                "Review": {"red": 1, "green": .98, "blue": .92},
                                "Invalid": {"red": 1, "green": .95, "blue": .95},
                            }.items()
                        ],
                    ]
                }
            )

        validity, validation_notes = assess_enquiry_validity(email, phone, message)
        sheet.append_row(
            [
                f"ENQ-{enquiry_id:05d}",
                format_india_datetime(created_at),
                name,
                email,
                phone or "Not provided",
                subject,
                message,
                "New",
                "Sent" if email_status == "sent" else "Saved only",
                "Portfolio Website",
                validity,
                validation_notes,
            ],
            value_input_option="USER_ENTERED",
            table_range="A:L",
        )
        return True


def send_enquiry_email(
    *,
    name: str,
    email: str,
    phone: str,
    subject: str,
    message: str,
) -> bool:
    if not all([MAIL_USERNAME, MAIL_PASSWORD, MAIL_RECEIVER]):
        app.logger.warning("Gmail settings are missing; enquiry saved to database only.")
        return False

    received_at = format_india_datetime()
    h_name = html_lib.escape(name)
    h_email = html_lib.escape(email)
    h_phone = html_lib.escape(phone or "Not provided")
    h_subject = html_lib.escape(subject)
    h_message = html_lib.escape(message)
    owner_mail = EmailMessage()
    owner_mail["Subject"] = f"New Portfolio Enquiry - {subject}"
    owner_mail["From"] = f"Shubham Chauhan Portfolio <{MAIL_USERNAME}>"
    owner_mail["To"] = MAIL_RECEIVER
    owner_mail["Reply-To"] = email
    owner_mail.set_content(
        f"""A new enquiry was submitted from the portfolio website.

CLIENT DETAILS
Name: {name}
Email: {email}
Phone: {phone or "Not provided"}
Subject: {subject}

MESSAGE
{message}

Received: {received_at}
"""
    )
    owner_mail.add_alternative(
        f"""<!doctype html>
<html><body style="margin:0;background:#f1f5f9;font-family:Arial,sans-serif;color:#172554">
<div style="max-width:640px;margin:28px auto;background:#fff;border-radius:18px;overflow:hidden;border:1px solid #dbeafe">
  <div style="padding:24px 28px;background:linear-gradient(135deg,#172554,#2563eb);color:#fff">
    <img src="cid:sc-owner-logo" alt="Shubham Chauhan" width="72" height="72" style="display:block;border-radius:14px;margin-bottom:15px;background:#fff">
    <div style="font-size:12px;letter-spacing:2px;font-weight:700">SHUBHAM CHAUHAN · SOFTWARE ENGINEER</div>
    <h1 style="margin:10px 0 0;font-size:24px">New portfolio enquiry</h1>
  </div>
  <div style="padding:28px">
    <p style="margin-top:0;color:#64748b">A new project enquiry has arrived. Reply directly to this email to contact the client.</p>
    <table role="presentation" style="width:100%;border-collapse:collapse;font-size:14px">
      <tr><td style="padding:9px 0;color:#64748b;width:145px">Client</td><td style="padding:9px 0;font-weight:700">{h_name}</td></tr>
      <tr><td style="padding:9px 0;color:#64748b">Email</td><td style="padding:9px 0">{h_email}</td></tr>
      <tr><td style="padding:9px 0;color:#64748b">Phone / WhatsApp</td><td style="padding:9px 0">{h_phone}</td></tr>
      <tr><td style="padding:9px 0;color:#64748b">Project</td><td style="padding:9px 0;font-weight:700">{h_subject}</td></tr>
    </table>
    <div style="margin-top:18px;padding:18px;border-radius:12px;background:#eff6ff;border-left:4px solid #2563eb;white-space:pre-wrap">{h_message}</div>
    <p style="margin:20px 0 0;color:#94a3b8;font-size:12px">Received: {received_at}</p>
  </div>
</div></body></html>""",
        subtype="html",
    )
    if LOGO_PATH.exists():
        owner_mail.get_payload()[-1].add_related(
            LOGO_PATH.read_bytes(),
            maintype="image",
            subtype="png",
            cid="<sc-owner-logo>",
            filename="shubham-chauhan-logo.png",
        )

    reply_mail = EmailMessage()
    reply_mail["Subject"] = f"Thank you for your enquiry - {subject}"
    reply_mail["From"] = f"Shubham Chauhan <{MAIL_USERNAME}>"
    reply_mail["To"] = email
    reply_mail["Reply-To"] = MAIL_RECEIVER
    reply_mail.set_content(
        f"""Hello {name},

Thank you for contacting Shubham Chauhan regarding {subject}.

Your enquiry has been received successfully. I will review your requirements and respond personally as soon as possible.

Your message:
{message}

Thanks,
Shubham Chauhan
Software Engineer
Mobile Apps | Web Development | Cloud & API Solutions
"""
    )
    reply_mail.add_alternative(
        f"""<!doctype html>
<html><body style="margin:0;background:#f1f5f9;font-family:Arial,sans-serif;color:#0f172a">
<div style="max-width:640px;margin:28px auto;background:#fff;border-radius:20px;overflow:hidden;border:1px solid #dbeafe;box-shadow:0 16px 40px rgba(15,23,42,.08)">
  <div style="padding:26px 30px;background:linear-gradient(135deg,#172554,#2563eb);color:#fff">
    <img src="cid:sc-client-logo" alt="Shubham Chauhan" width="82" height="82" style="display:block;border-radius:16px;margin-bottom:16px;background:#fff">
    <div style="font-size:12px;letter-spacing:2px;font-weight:700">SC · SHUBHAM CHAUHAN</div>
    <h1 style="margin:11px 0 5px;font-size:25px">Thank you for reaching out.</h1>
    <p style="margin:0;color:#dbeafe;font-size:14px">Your project enquiry has been received successfully.</p>
  </div>
  <div style="padding:30px">
    <p style="font-size:16px">Hello <strong>{h_name}</strong>,</p>
    <p style="line-height:1.7;color:#475569">Thank you for contacting me regarding <strong>{h_subject}</strong>. I will review the requirements and respond personally as soon as possible.</p>
    <div style="margin:22px 0;padding:18px;border-radius:14px;background:#eff6ff;border:1px solid #dbeafe">
      <div style="margin-bottom:8px;color:#2563eb;font-size:12px;font-weight:700;letter-spacing:1px">YOUR ENQUIRY</div>
      <div style="white-space:pre-wrap;line-height:1.65;color:#334155">{h_message}</div>
    </div>
    <p style="line-height:1.6">Thanks,<br><strong style="color:#172554">Shubham Chauhan</strong><br><span style="color:#64748b">Software Engineer · Mobile Apps · Web Development · Cloud & API Solutions</span></p>
  </div>
  <div style="padding:16px 30px;background:#f8fafc;color:#94a3b8;font-size:11px;text-align:center">This is an automatic confirmation from the Shubham Chauhan portfolio.</div>
</div></body></html>""",
        subtype="html",
    )
    if LOGO_PATH.exists():
        reply_mail.get_payload()[-1].add_related(
            LOGO_PATH.read_bytes(),
            maintype="image",
            subtype="png",
            cid="<sc-client-logo>",
            filename="shubham-chauhan-logo.png",
        )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as smtp:
        smtp.login(MAIL_USERNAME, MAIL_PASSWORD)
        smtp.send_message(owner_mail)
        try:
            smtp.send_message(reply_mail)
        except Exception:
            app.logger.exception("Owner notification sent, but client auto-reply failed.")

    return True


def local_assistant_reply(message: str) -> str:
    """Useful fallback when no OpenAI key is configured or the API is unavailable."""
    lowered = message.lower()

    if any(term in lowered for term in ("e-com", "ecom", "ecommerce", "e-commerce", "shop")):
        return (
            "Shubham can build a single-vendor or multi-vendor e-commerce platform "
            "with Flutter apps, a responsive website, product management, cart, "
            "payments, order tracking, inventory and an admin dashboard. "
            "Do you need a single store or a marketplace with multiple sellers?"
        )

    if any(term in lowered for term in ("dating", "matchmaking", "matrimonial")):
        return (
            "A dating platform can include profile verification, preferences, swipe "
            "matching, chat, subscriptions, package limits, payments, report/block "
            "controls and an admin moderation panel. Which packages and platforms "
            "do you want: Android, iOS, website, or all three?"
        )

    if any(term in lowered for term in ("school", "erp", "inventory", "crm")):
        return (
            "Shubham develops role-based ERP systems with dashboards, permissions, "
            "reports, notifications, mobile apps and Python APIs. Tell me the user "
            "roles and the most important modules for your system."
        )

    if any(term in lowered for term in ("flutter", "android", "ios", "mobile app")):
        return (
            "Shubham builds Flutter apps for Android and iOS with authentication, "
            "Firebase, REST APIs, payments, maps, notifications and admin panels. "
            "Please describe your main user journey."
        )

    if any(term in lowered for term in ("website", "portfolio", "web app")):
        return (
            "Shubham can build a responsive website or web application with Python, "
            "Flask or FastAPI, dynamic content, database, admin panel, SEO and enquiry "
            "automation. What type of business is the website for?"
        )

    return (
        "I can help you plan mobile apps, websites, e-commerce platforms, ERP systems, "
        "Python APIs and admin dashboards. Share your project idea, required platforms "
        "and key features, and I will structure the solution."
    )


def sanitise_history(history: Any) -> list[dict[str, str]]:
    clean: list[dict[str, str]] = []
    if not isinstance(history, list):
        return clean

    for item in history[-14:]:
        if not isinstance(item, dict):
            continue
        role = item.get("role")
        content = safe_text(item.get("content"), 4000)
        if role in {"user", "assistant"} and content:
            clean.append({"role": role, "content": content})
    return clean


ALLOWED_CHAT_FILES = {
    "image/jpeg": ".jpg",
    "image/png": ".png",
    "image/webp": ".webp",
    "application/pdf": ".pdf",
}


def read_chat_attachments() -> tuple[list[dict[str, Any]], str | None]:
    attachments: list[dict[str, Any]] = []
    total_size = 0
    for uploaded in request.files.getlist("attachments")[:3]:
        content_type = (uploaded.mimetype or "").lower()
        if content_type not in ALLOWED_CHAT_FILES:
            return [], "Only JPG, PNG, WebP and PDF files are supported."
        data = uploaded.read(5 * 1024 * 1024 + 1)
        if len(data) > 5 * 1024 * 1024:
            return [], "Each attachment must be 5 MB or smaller."
        total_size += len(data)
        if total_size > 10 * 1024 * 1024:
            return [], "Attachments must be 10 MB or smaller in total."
        attachments.append(
            {
                "name": safe_text(uploaded.filename, 180) or f"attachment{ALLOWED_CHAT_FILES[content_type]}",
                "content_type": content_type,
                "data": data,
            }
        )
    return attachments, None


def send_chat_copy(message: str, attachments: list[dict[str, Any]]) -> None:
    """Forward each visitor chat message and its files to the portfolio owner."""
    if not all([MAIL_USERNAME, MAIL_PASSWORD, MAIL_RECEIVER]):
        return
    mail = EmailMessage()
    mail["Subject"] = "Portfolio AI Chat - New visitor message"
    mail["From"] = f"Shubham Chauhan Portfolio <{MAIL_USERNAME}>"
    mail["To"] = MAIL_RECEIVER
    file_names = ", ".join(item["name"] for item in attachments) or "None"
    mail.set_content(
        f"Visitor message:\n\n{message or 'Attachment shared without text'}\n\n"
        f"Attachments: {file_names}\nReceived: {format_india_datetime()}"
    )
    mail.add_alternative(
        f"""<!doctype html><html><body style="margin:0;background:#f1f5f9;font-family:Arial,sans-serif;color:#0f172a">
<div style="max-width:620px;margin:28px auto;background:#fff;border:1px solid #dbeafe;border-radius:18px;overflow:hidden">
<div style="padding:24px;background:linear-gradient(135deg,#172554,#2563eb);color:#fff">
<img src="cid:sc-chat-logo" width="72" height="72" alt="Shubham Chauhan" style="display:block;border-radius:14px;background:#fff;margin-bottom:14px">
<strong style="font-size:20px">New AI chat message</strong></div>
<div style="padding:26px"><p style="white-space:pre-wrap;line-height:1.7">{html_lib.escape(message or 'Attachment shared without text')}</p>
<p style="padding:14px;background:#eff6ff;border-radius:12px"><strong>Attachments:</strong> {html_lib.escape(file_names)}</p>
<p style="color:#64748b;font-size:12px">Received: {format_india_datetime()}</p></div></div></body></html>""",
        subtype="html",
    )
    if LOGO_PATH.exists():
        mail.get_payload()[-1].add_related(
            LOGO_PATH.read_bytes(), maintype="image", subtype="png",
            cid="<sc-chat-logo>", filename="shubham-chauhan-logo.png",
        )
    for item in attachments:
        maintype, subtype = item["content_type"].split("/", 1)
        mail.add_attachment(
            item["data"],
            maintype=maintype,
            subtype=subtype,
            filename=item["name"],
        )
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as smtp:
        smtp.login(MAIL_USERNAME, MAIL_PASSWORD)
        smtp.send_message(mail)


def forward_chat_copy_safely(message: str, attachments: list[dict[str, Any]]) -> None:
    try:
        send_chat_copy(message, attachments)
    except Exception:
        app.logger.exception("Unable to forward the visitor chat message.")


@app.context_processor
def global_template_values() -> dict[str, Any]:
    whatsapp_message = (
        "Hello Shubham, I visited your portfolio and want to discuss a project."
    )
    whatsapp_url = (
        f"https://wa.me/{WHATSAPP_NUMBER}"
        f"?text={whatsapp_message.replace(' ', '%20')}"
    )
    return {
        "current_year": datetime.now().year,
        "whatsapp_url": whatsapp_url,
        "ai_enabled": bool(openai_client),
    }


@app.get("/")
def home() -> str:
    projects = [
        {
            "icon": "fa-layer-group",
            "category": "DIGITAL PLATFORM",
            "title": "WASEF",
            "image": "images/projects/wasef-logo.jpeg",
            "description": "A structured digital product focused on a clear, reliable and responsive user experience.",
            "tags": ["Product", "Mobile", "API"],
        },
        {
            "icon": "fa-chart-line",
            "category": "BUSINESS OPERATIONS",
            "title": "Biz-Control",
            "image": "images/projects/biz-control-logo.jpeg",
            "description": "A business-control solution designed to organise day-to-day operations and management workflows.",
            "tags": ["Business", "Dashboard", "Workflow"],
            "store_live": False,
        },
        {
            "icon": "fa-motorcycle",
            "category": "COLLECTION PLATFORM",
            "title": "CollectX Agency",
            "image": "images/projects/collectx-agency-icon.png",
            "description": "The agency-facing CollectX application for coordinated collection and operational workflows.",
            "tags": ["Agency App", "Collections", "Operations"],
            "store_url": "https://play.google.com/store/apps/details?id=com.arramton.collectx&hl=en_IN",
            "store_live": True,
        },
        {
            "icon": "fa-motorcycle",
            "category": "COLLECTION PLATFORM",
            "title": "CollectX Rider",
            "image": "images/projects/collectx-rider-icon.png",
            "description": "The rider application supporting assigned collection tasks and field operations.",
            "tags": ["Rider App", "Field Work", "Operations"],
            "store_url": "https://play.google.com/store/apps/details?id=com.arramton.collectxrider&hl=en_IN",
            "store_live": True,
        },
        {
            "icon": "fa-shirt",
            "category": "CLOTHING STORAGE",
            "title": "Closet",
            "description": "A customer application designed to organise wardrobe storage and clothing-management requests.",
            "tags": ["User App", "Wardrobe", "Storage"],
            "store_live": False,
        },
        {
            "icon": "fa-truck-fast",
            "category": "CLOTHING STORAGE",
            "title": "Closet Rider",
            "description": "A field-team application for clothing pickups, task updates and operational coordination.",
            "tags": ["Rider App", "Pickups", "Operations"],
            "store_live": False,
        },
        {
            "icon": "fa-cake-candles",
            "category": "FOOD & COMMERCE",
            "title": "Cakingom - Online Cake Delivery",
            "image": "images/projects/cakingom-icon.png",
            "description": "Cake, flowers, personalised gifts and celebration-decoration ordering with doorstep delivery.",
            "tags": ["Shopping", "Orders", "Delivery"],
            "store_url": "https://play.google.com/store/apps/details?id=com.arramton.cakingom",
            "store_live": True,
        },
        {
            "icon": "fa-store",
            "category": "BUSINESS",
            "title": "Keyenzy",
            "image": "images/projects/keyenzy-icon.png",
            "description": "A mobile business application that helps registered shop owners manage their online presence and operations.",
            "tags": ["Shop Owners", "Business", "Management"],
            "store_url": "https://play.google.com/store/apps/details?id=com.arramton.keyenzy",
            "store_live": True,
        },
        {
            "icon": "fa-car-side",
            "category": "AUTO MARKETPLACE",
            "title": "Carenzy",
            "image": "images/projects/carenzy-icon.png",
            "description": "A used-car marketplace connecting buyers, sellers and dealers through listings, bidding and lead workflows.",
            "tags": ["Marketplace", "Bidding", "Automotive"],
            "store_url": "https://play.google.com/store/apps/details?id=com.car.carenzy",
            "store_live": True,
        },
        {
            "icon": "fa-key",
            "category": "BUSINESS ORDERING",
            "title": "KMD Keys",
            "image": "images/projects/kmd-keys-icon.png",
            "description": "A specialised ordering application created for keymakers to submit orders for faster delivery.",
            "tags": ["Orders", "Keymakers", "Delivery"],
            "store_url": "https://play.google.com/store/apps/details?id=com.app.kmdkeys",
            "store_live": True,
        },
        {
            "icon": "fa-print",
            "category": "PRINT SERVICES",
            "title": "ZoomPrint",
            "description": "A mobile product centred on convenient digital print-service requests and workflow coordination.",
            "tags": ["Printing", "Orders", "Services"],
            "store_live": False,
        },
        {
            "icon": "fa-house-circle-check",
            "category": "HOME SERVICES",
            "title": "The NinjaCare",
            "image": "images/projects/ninjacare-user-icon.png",
            "description": "A customer app for scheduling household cleaning, cooking and domestic-helper services.",
            "tags": ["User App", "Bookings", "Home Services"],
            "store_live": False,
        },
        {
            "icon": "fa-people-carry-box",
            "category": "HOME SERVICES",
            "title": "NinjaCare Partner",
            "image": "images/projects/ninjacare-partner-icon.png",
            "description": "The partner-side experience supporting assigned household-service work and delivery operations.",
            "tags": ["Partner App", "Services", "Operations"],
            "store_live": False,
        },
        {
            "icon": "fa-burger",
            "category": "FOOD DELIVERY",
            "title": "Food Go",
            "description": "A food-ordering product focused on a quick customer journey and organised fulfilment workflow.",
            "tags": ["Food", "Delivery", "Orders"],
        },
        {
            "icon": "fa-plane-departure",
            "category": "TRAVEL",
            "title": "Flight Booking",
            "description": "A responsive flight-search and booking experience with a clear, traveller-friendly journey.",
            "tags": ["Travel", "Booking", "API"],
        },
        {
            "icon": "fa-user-graduate",
            "category": "EDUCATION MANAGEMENT",
            "title": "Student Management",
            "description": "A structured student-management system for organising academic and administrative information.",
            "tags": ["Education", "Management", "Dashboard"],
        },
        {
            "icon": "fa-screwdriver-wrench",
            "category": "SERVICE PLATFORM",
            "title": "Metri Serv",
            "description": "A service-oriented digital product designed to connect requests with an efficient operational workflow.",
            "tags": ["Services", "Operations", "Platform"],
        },
    ]

    reviews = [
        {
            "company": "WASEF",
            "image": "images/projects/wasef-logo.jpeg",
            "rating": 5,
            "text": "Shubham understood our requirements and presented the WASEF project in a clean, professional way. He paid attention to our feedback, made the requested improvements carefully, and produced work that matched the direction we had discussed.",
        },
        {
            "company": "Biz-Control",
            "image": "images/projects/biz-control-logo.jpeg",
            "rating": 4,
            "text": "Our experience with Shubham was positive and straightforward. He listened to the needs of Biz-Control, organised the work clearly, and created a practical final result. The quality of the completed project was very good.",
        },
        {
            "company": "CollectX Agency & Rider",
            "image": "images/projects/collectx-agency-icon.png",
            "rating": 5,
            "text": "The CollectX work was handled with care from the initial discussion through the final updates. Shubham remained cooperative, understood the separate agency and rider requirements, and completed the project in a polished and well-organised manner.",
        },
        {
            "company": "Keyenzy",
            "image": "images/projects/keyenzy-icon.png",
            "rating": 4,
            "text": "We appreciated Shubham’s patient and supportive approach while working on Keyenzy. Our ideas were considered properly, changes were addressed without confusion, and the finished application looks clean, useful, and professionally prepared.",
        },
        {
            "company": "Cakingom",
            "image": "images/projects/cakingom-icon.png",
            "rating": 5,
            "text": "Working with Shubham on Cakingom felt smooth and comfortable. He communicated clearly, gave proper attention to the details we shared, and refined the project based on our suggestions. We are happy with the overall outcome.",
        },
        {
            "company": "The NinjaCare",
            "image": "images/projects/ninjacare-user-icon.png",
            "rating": 5,
            "text": "Good work was delivered for The NinjaCare project. Shubham took time to understand what was needed for the customer and partner sides, responded well to feedback, and ensured the final presentation was clear and consistent.",
        },
    ]

    return render_template("index.html", projects=projects, reviews=reviews)


@app.post("/contact")
def contact() -> Response:
    if safe_text(request.form.get("company_website"), 200):
        flash("Thank you. Your enquiry was received successfully.", "success")
        return redirect(url_for("home", _anchor="contact"))

    name = safe_text(request.form.get("name"), 100)
    email = safe_text(request.form.get("email"), 180)
    phone = safe_text(request.form.get("phone"), 40)
    subject = safe_text(request.form.get("subject"), 160)
    message = safe_text(request.form.get("message"), 5000)

    if not all([name, email, subject, message]):
        flash("Please complete all required fields.", "error")
        return redirect(url_for("home", _anchor="contact"))

    if not looks_like_email(email):
        flash("Please enter a valid email address.", "error")
        return redirect(url_for("home", _anchor="contact"))

    created_at = datetime.now(INDIA_TIMEZONE).isoformat(timespec="seconds")

    with database_connection() as connection:
        cursor = connection.execute(
            """
            INSERT INTO contact_messages
            (name, email, phone, subject, message, email_status, created_at)
            VALUES (?, ?, ?, ?, ?, 'pending', ?)
            """,
            (name, email, phone, subject, message, created_at),
        )
        message_id = cursor.lastrowid
        connection.commit()

    email_sent = False
    try:
        email_sent = send_enquiry_email(
            name=name,
            email=email,
            phone=phone,
            subject=subject,
            message=message,
        )
    except smtplib.SMTPAuthenticationError:
        app.logger.exception("Gmail authentication failed.")
    except Exception:
        app.logger.exception("Unable to deliver enquiry email.")

    with database_connection() as connection:
        connection.execute(
            "UPDATE contact_messages SET email_status = ? WHERE id = ?",
            ("sent" if email_sent else "saved_only", message_id),
        )
        connection.commit()

    try:
        append_enquiry_to_excel(
            enquiry_id=int(message_id),
            created_at=created_at,
            name=name,
            email=email,
            phone=phone,
            subject=subject,
            message=message,
            email_status="sent" if email_sent else "saved_only",
        )
    except Exception:
        app.logger.exception("Unable to update the Excel enquiry register.")

    try:
        append_enquiry_to_google_sheet(
            enquiry_id=int(message_id),
            created_at=created_at,
            name=name,
            email=email,
            phone=phone,
            subject=subject,
            message=message,
            email_status="sent" if email_sent else "saved_only",
        )
    except Exception:
        app.logger.exception("Unable to update the Google Sheet enquiry register.")

    if email_sent:
        flash("Thank you. Your enquiry was sent successfully.", "success")
    else:
        flash(
            "Your enquiry was saved. Email delivery is not configured yet, "
            "so you can also contact me on WhatsApp.",
            "warning",
        )

    return redirect(url_for("home", _anchor="contact"))


@app.post("/api/chat-stream")
def chat_stream() -> Response:
    if request.content_type and request.content_type.startswith("multipart/form-data"):
        message = safe_text(request.form.get("message"), 4000)
        try:
            raw_history = json.loads(request.form.get("history", "[]"))
        except json.JSONDecodeError:
            raw_history = []
        attachments, attachment_error = read_chat_attachments()
        if attachment_error:
            return Response(json.dumps({"error": attachment_error}), status=400, content_type="application/json")
    else:
        body = request.get_json(silent=True) or {}
        message = safe_text(body.get("message"), 4000)
        raw_history = body.get("history")
        attachments = []
    history = sanitise_history(raw_history)

    if not message and not attachments:
        return Response(
            json.dumps({"error": "Message is required."}),
            status=400,
            content_type="application/json",
        )

    def send_event(payload: dict[str, Any]) -> str:
        return f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"

    @stream_with_context
    def generate():
        threading.Thread(
            target=forward_chat_copy_safely,
            args=(message, attachments),
            daemon=True,
        ).start()

        if openai_client is None:
            fallback_message = message or "The visitor shared a project attachment."
            yield send_event({"type": "delta", "text": local_assistant_reply(fallback_message)})
            yield send_event({"type": "done"})
            return

        user_content: list[dict[str, Any]] = [
            {"type": "input_text", "text": message or "Please review the attached project file."}
        ]
        for item in attachments:
            encoded = base64.b64encode(item["data"]).decode("ascii")
            data_url = f"data:{item['content_type']};base64,{encoded}"
            if item["content_type"].startswith("image/"):
                user_content.append({"type": "input_image", "image_url": data_url})
            else:
                user_content.append({"type": "input_file", "filename": item["name"], "file_data": data_url})
        conversation = [*history, {"role": "user", "content": user_content}]

        try:
            stream = openai_client.responses.create(
                model=OPENAI_MODEL,
                instructions=(
                    "You are Shubham Chauhan's professional portfolio AI assistant. "
                    "Shubham provides Python, Flask, FastAPI, Flutter, Android, iOS, "
                    "ERP, e-commerce, business website, desktop software, API and admin "
                    "dashboard development. Understand abbreviations and spelling errors. "
                    "Carefully analyse attached images and PDFs when present. Help the visitor "
                    "clarify requirements and recommend relevant modules using only supplied "
                    "information and the capabilities listed here. Explicitly say when information "
                    "is uncertain or needs human confirmation. "
                    "Never invent verified clients, fixed pricing, guaranteed timelines, "
                    "credentials or capabilities not stated here. Ask no more than one "
                    "useful follow-up question per response. Do not add a signature because the "
                    "application adds it automatically. Keep replies professional, clear, friendly "
                    "and normally below 220 words."
                ),
                input=conversation,
                stream=True,
            )

            emitted = False
            for event in stream:
                event_type = getattr(event, "type", "")
                if event_type == "response.output_text.delta":
                    delta = getattr(event, "delta", "")
                    if delta:
                        emitted = True
                        yield send_event({"type": "delta", "text": delta})
                elif event_type == "error":
                    raise RuntimeError(str(event))

            if not emitted:
                yield send_event(
                    {
                        "type": "delta",
                        "text": "I could not generate a reply. Please try again.",
                    }
                )
            yield send_event({"type": "done"})

        except Exception:
            app.logger.exception("OpenAI streaming request failed.")
            yield send_event(
                {
                    "type": "notice",
                    "text": "AI service is temporarily unavailable; showing a local assistant reply.",
                }
            )
            fallback_message = message or "The visitor shared a project attachment."
            yield send_event({"type": "delta", "text": local_assistant_reply(fallback_message)})
            yield send_event({"type": "done"})

    return Response(
        generate(),
        content_type="text/event-stream; charset=utf-8",
        headers={
            "Cache-Control": "no-cache, no-transform",
            "X-Accel-Buffering": "no",
            "Connection": "keep-alive",
        },
    )


@app.get("/health")
def health() -> dict[str, Any]:
    return {
        "status": "ok",
        "ai_configured": bool(openai_client),
        "email_configured": bool(
            MAIL_USERNAME and MAIL_PASSWORD and MAIL_RECEIVER
        ),
        "google_sheets_configured": bool(
            GOOGLE_SHEET_ID and GOOGLE_SERVICE_ACCOUNT_PATH.is_file()
        ),
    }


@app.after_request
def add_security_headers(response: Response) -> Response:
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    response.headers.setdefault("Permissions-Policy", "camera=(), microphone=(), geolocation=()")
    return response


if __name__ == "__main__":
    initialise_database()
    debug_enabled = os.getenv("FLASK_DEBUG", "false").lower() == "true"
    app.run(host="127.0.0.1", port=5000, debug=debug_enabled, threaded=True)
