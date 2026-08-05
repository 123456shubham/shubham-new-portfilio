from __future__ import annotations

from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path
from typing import Any
import json
import os
import re
import smtplib
import sqlite3

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
    MAX_CONTENT_LENGTH=32 * 1024,
)

MAIL_USERNAME = os.getenv("MAIL_USERNAME", "").strip()
MAIL_PASSWORD = os.getenv("MAIL_PASSWORD", "").replace(" ", "").strip()
MAIL_RECEIVER = os.getenv("MAIL_RECEIVER", MAIL_USERNAME).strip()

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "").strip()
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-5-mini").strip()

WHATSAPP_NUMBER = re.sub(
    r"\D",
    "",
    os.getenv("WHATSAPP_NUMBER", "919999999999"),
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

    mail = EmailMessage()
    mail["Subject"] = f"New Portfolio Enquiry — {subject}"
    mail["From"] = MAIL_USERNAME
    mail["To"] = MAIL_RECEIVER
    mail["Reply-To"] = email
    mail.set_content(
        f"""A new enquiry was submitted from the portfolio website.

CLIENT DETAILS
Name: {name}
Email: {email}
Phone: {phone or "Not provided"}
Subject: {subject}

MESSAGE
{message}

Received: {datetime.now(timezone.utc).isoformat(timespec="seconds")}
"""
    )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as smtp:
        smtp.login(MAIL_USERNAME, MAIL_PASSWORD)
        smtp.send_message(mail)

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
            "icon": "fa-school",
            "category": "ERP PLATFORM",
            "title": "School Management ERP",
            "description": "Role-based mobile apps, APIs, administration and operational modules.",
            "tags": ["Flutter", "FastAPI", "PostgreSQL"],
        },
        {
            "icon": "fa-cart-shopping",
            "category": "COMMERCE",
            "title": "E-commerce Platform",
            "description": "Customer experience, payments, inventory, order tracking and admin tools.",
            "tags": ["Python", "Flutter", "Payments"],
        },
        {
            "icon": "fa-industry",
            "category": "BUSINESS SOFTWARE",
            "title": "Inventory Desktop Suite",
            "description": "Operational desktop software for inventory, reports and workflow automation.",
            "tags": ["PySide6", "SQLite", "Reports"],
        },
        {
            "icon": "fa-wind",
            "category": "CORPORATE WEB",
            "title": "HVAC Product Website",
            "description": "Responsive industrial product presentation with enquiries and animations.",
            "tags": ["Flask", "JavaScript", "SEO"],
        },
    ]

    reviews = [
        {
            "initials": "AV",
            "name": "Amit Verma",
            "position": "Operations Manager",
            "text": "The workflow was understood quickly and converted into a clean, practical system. Communication and revisions were handled professionally.",
        },
        {
            "initials": "NS",
            "name": "Neha Sharma",
            "position": "School Administrator",
            "text": "The connected modules were structured clearly, and the interface was much easier for our staff to operate.",
        },
        {
            "initials": "RM",
            "name": "Rohit Mehta",
            "position": "Business Owner",
            "text": "The website looked professional on desktop and mobile, and the final code was organised and maintainable.",
        },
    ]

    return render_template("index.html", projects=projects, reviews=reviews)


@app.post("/contact")
def contact() -> Response:
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

    created_at = datetime.now(timezone.utc).isoformat(timespec="seconds")

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
    body = request.get_json(silent=True) or {}
    message = safe_text(body.get("message"), 4000)
    history = sanitise_history(body.get("history"))

    if not message:
        return Response(
            json.dumps({"error": "Message is required."}),
            status=400,
            content_type="application/json",
        )

    def send_event(payload: dict[str, Any]) -> str:
        return f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"

    @stream_with_context
    def generate():
        if openai_client is None:
            yield send_event({"type": "delta", "text": local_assistant_reply(message)})
            yield send_event({"type": "done"})
            return

        conversation = [*history, {"role": "user", "content": message}]

        try:
            stream = openai_client.responses.create(
                model=OPENAI_MODEL,
                instructions=(
                    "You are Shubham Chauhan's professional portfolio AI assistant. "
                    "Shubham provides Python, Flask, FastAPI, Flutter, Android, iOS, "
                    "ERP, e-commerce, business website, desktop software, API and admin "
                    "dashboard development. Understand abbreviations and spelling errors. "
                    "Help the visitor clarify requirements and recommend relevant modules. "
                    "Never invent verified clients, fixed pricing, guaranteed timelines, "
                    "credentials or capabilities not stated here. Ask no more than one "
                    "useful follow-up question per response. Keep replies clear, friendly "
                    "and normally below 180 words."
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
            yield send_event({"type": "delta", "text": local_assistant_reply(message)})
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
    }


if __name__ == "__main__":
    initialise_database()
    app.run(host="127.0.0.1", port=5000, debug=True, threaded=True)
