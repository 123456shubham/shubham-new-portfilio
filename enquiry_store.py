from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from typing import Any, Iterator
import os
import uuid

import psycopg
from psycopg.rows import dict_row


DATABASE_URL = os.getenv("DATABASE_URL", "").strip()


def configured() -> bool:
    return bool(DATABASE_URL)


@contextmanager
def connection() -> Iterator[psycopg.Connection]:
    if not DATABASE_URL:
        raise RuntimeError("DATABASE_URL is not configured")
    with psycopg.connect(DATABASE_URL, row_factory=dict_row) as conn:
        yield conn


def initialise() -> None:
    with connection() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS enquiries (
                enquiry_id TEXT PRIMARY KEY,
                submission_token UUID UNIQUE,
                received_at TIMESTAMPTZ NOT NULL,
                client_name TEXT NOT NULL,
                email TEXT NOT NULL,
                phone TEXT NOT NULL DEFAULT '',
                project_subject TEXT NOT NULL,
                project_details TEXT NOT NULL,
                project_amount NUMERIC(14, 2),
                amount_received NUMERIC(14, 2) NOT NULL DEFAULT 0,
                payment_date DATE,
                payment_status TEXT NOT NULL DEFAULT 'Pending',
                lead_status TEXT NOT NULL DEFAULT 'New',
                email_status TEXT NOT NULL DEFAULT 'Pending',
                source TEXT NOT NULL DEFAULT 'Portfolio Website',
                validity TEXT NOT NULL DEFAULT 'Valid',
                validation_notes TEXT NOT NULL DEFAULT '',
                updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                sheet_row INTEGER,
                sync_status TEXT NOT NULL DEFAULT 'Pending',
                last_synced_at TIMESTAMPTZ,
                quotation_number TEXT,
                quotation_amount NUMERIC(14, 2),
                quotation_currency TEXT NOT NULL DEFAULT 'INR',
                delivery_days INTEGER,
                quotation_valid_until DATE,
                payment_terms TEXT NOT NULL DEFAULT '',
                quotation_notes TEXT NOT NULL DEFAULT '',
                quotation_status TEXT NOT NULL DEFAULT 'Pending',
                quotation_sent_at TIMESTAMPTZ,
                payment_due_date DATE,
                payment_email_status TEXT NOT NULL DEFAULT 'Pending',
                payment_receipt_sent_at TIMESTAMPTZ,
                payment_reminder_sent_at TIMESTAMPTZ
            )
            """
        )
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS amount_received NUMERIC(14, 2) NOT NULL DEFAULT 0")
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS payment_date DATE")
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS payment_status TEXT NOT NULL DEFAULT 'Pending'")
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS quotation_number TEXT")
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS quotation_amount NUMERIC(14, 2)")
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS quotation_currency TEXT NOT NULL DEFAULT 'INR'")
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS delivery_days INTEGER")
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS quotation_valid_until DATE")
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS payment_terms TEXT NOT NULL DEFAULT ''")
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS quotation_notes TEXT NOT NULL DEFAULT ''")
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS quotation_status TEXT NOT NULL DEFAULT 'Pending'")
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS quotation_sent_at TIMESTAMPTZ")
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS payment_due_date DATE")
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS payment_email_status TEXT NOT NULL DEFAULT 'Pending'")
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS payment_receipt_sent_at TIMESTAMPTZ")
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS payment_reminder_sent_at TIMESTAMPTZ")
        conn.execute(
            "CREATE INDEX IF NOT EXISTS enquiries_updated_at_idx ON enquiries(updated_at DESC)"
        )


def parse_amount(value: Any) -> Decimal | None:
    text = str(value or "").strip().lstrip("'")
    text = text.replace(",", "").replace("₹", "").replace("INR", "").strip()
    if not text:
        return None
    try:
        amount = Decimal(text).quantize(Decimal("0.01"))
    except InvalidOperation as exc:
        raise ValueError("Project amount must be a valid number") from exc
    if amount < 0 or amount > Decimal("999999999999.99"):
        raise ValueError("Project amount is outside the supported range")
    return amount


def create_enquiry(data: dict[str, Any]) -> tuple[dict[str, Any], bool]:
    initialise()
    token = uuid.UUID(str(data["submission_token"]))
    enquiry_id = str(data.get("enquiry_id") or f"ENQ-{uuid.uuid4().hex[:12].upper()}")
    with connection() as conn:
        row = conn.execute(
            """
            INSERT INTO enquiries (
                enquiry_id, submission_token, received_at, client_name, email,
                phone, project_subject, project_details, project_amount,
                lead_status, email_status, source, validity, validation_notes,
                quotation_number, quotation_amount, quotation_currency,
                delivery_days, quotation_valid_until, payment_terms,
                quotation_notes, quotation_status
            ) VALUES (
                %(enquiry_id)s, %(submission_token)s, %(received_at)s, %(client_name)s,
                %(email)s, %(phone)s, %(project_subject)s, %(project_details)s,
                %(project_amount)s, 'New', 'Pending', 'Portfolio Website',
                %(validity)s, %(validation_notes)s, %(quotation_number)s,
                %(quotation_amount)s, %(quotation_currency)s, %(delivery_days)s,
                %(quotation_valid_until)s, %(payment_terms)s,
                %(quotation_notes)s, 'Pending'
            )
            ON CONFLICT (submission_token) DO NOTHING
            RETURNING *
            """,
            {**data, "enquiry_id": enquiry_id, "submission_token": token},
        ).fetchone()
        if row:
            return dict(row), True
        existing = conn.execute(
            "SELECT * FROM enquiries WHERE submission_token = %s", (token,)
        ).fetchone()
        return dict(existing), False


def update_delivery(enquiry_id: str, email_status: str) -> None:
    with connection() as conn:
        conn.execute(
            "UPDATE enquiries SET email_status=%s, updated_at=NOW() WHERE enquiry_id=%s",
            (email_status, enquiry_id),
        )


def update_quotation_delivery(enquiry_id: str, status: str) -> None:
    with connection() as conn:
        conn.execute(
            """UPDATE enquiries SET quotation_status=%s,
               quotation_sent_at=CASE WHEN %s='Sent' THEN NOW() ELSE quotation_sent_at END,
               lead_status=CASE WHEN %s='Sent' THEN 'Proposal Sent' ELSE lead_status END,
               updated_at=NOW() WHERE enquiry_id=%s""",
            (status, status, status, enquiry_id),
        )


def mark_payment_receipt_sent(enquiry_id: str) -> None:
    with connection() as conn:
        conn.execute(
            """UPDATE enquiries SET payment_email_status='Receipt Sent',
               payment_receipt_sent_at=NOW(), updated_at=NOW()
               WHERE enquiry_id=%s""",
            (enquiry_id,),
        )


def mark_payment_reminder_sent(enquiry_id: str) -> None:
    with connection() as conn:
        conn.execute(
            """UPDATE enquiries SET payment_email_status='Reminder Sent',
               payment_reminder_sent_at=NOW(), updated_at=NOW()
               WHERE enquiry_id=%s""",
            (enquiry_id,),
        )


def reset_payment_reminder(enquiry_id: str) -> None:
    """Allow one fresh automatic reminder after the due date is edited."""
    with connection() as conn:
        conn.execute(
            """UPDATE enquiries SET payment_reminder_sent_at=NULL,
               payment_email_status=CASE
                   WHEN payment_email_status='Reminder Sent' THEN 'Pending'
                   ELSE payment_email_status
               END,
               updated_at=NOW() WHERE enquiry_id=%s""",
            (enquiry_id,),
        )


def mark_synced(enquiry_id: str, sheet_row: int) -> None:
    with connection() as conn:
        conn.execute(
            """UPDATE enquiries SET sheet_row=%s, sync_status='Synced',
               last_synced_at=NOW() WHERE enquiry_id=%s""",
            (sheet_row, enquiry_id),
        )


def upsert_from_sheet(data: dict[str, Any], sheet_row: int) -> dict[str, Any]:
    initialise()
    enquiry_id = str(data.get("enquiry_id") or f"ENQ-{uuid.uuid4().hex[:12].upper()}")
    received_at = data.get("received_at") or datetime.now(timezone.utc)
    with connection() as conn:
        row = conn.execute(
            """
            INSERT INTO enquiries (
                enquiry_id, received_at, client_name, email, phone, project_subject,
                project_details, project_amount, amount_received, payment_date,
                payment_status, lead_status, email_status, source,
                validity, validation_notes, updated_at, sheet_row, sync_status, last_synced_at
                , quotation_number, quotation_amount, quotation_currency, delivery_days,
                quotation_valid_until, payment_terms, quotation_notes, quotation_status,
                quotation_sent_at, payment_due_date, payment_email_status,
                payment_receipt_sent_at, payment_reminder_sent_at
            ) VALUES (
                %(enquiry_id)s, %(received_at)s, %(client_name)s, %(email)s, %(phone)s,
                %(project_subject)s, %(project_details)s, %(project_amount)s,
                %(amount_received)s, %(payment_date)s, %(payment_status)s,
                %(lead_status)s, %(email_status)s, %(source)s, %(validity)s,
                %(validation_notes)s, NOW(), %(sheet_row)s, 'Synced', NOW(),
                %(quotation_number)s, %(quotation_amount)s, %(quotation_currency)s,
                %(delivery_days)s, %(quotation_valid_until)s, %(payment_terms)s,
                %(quotation_notes)s, %(quotation_status)s, %(quotation_sent_at)s
                , %(payment_due_date)s, %(payment_email_status)s,
                %(payment_receipt_sent_at)s, %(payment_reminder_sent_at)s
            )
            ON CONFLICT (enquiry_id) DO UPDATE SET
                client_name=EXCLUDED.client_name, email=EXCLUDED.email,
                phone=EXCLUDED.phone, project_subject=EXCLUDED.project_subject,
                project_details=EXCLUDED.project_details,
                project_amount=EXCLUDED.project_amount,
                amount_received=EXCLUDED.amount_received,
                payment_date=EXCLUDED.payment_date,
                payment_status=EXCLUDED.payment_status,
                lead_status=EXCLUDED.lead_status,
                email_status=EXCLUDED.email_status, source=EXCLUDED.source,
                validity=EXCLUDED.validity, validation_notes=EXCLUDED.validation_notes,
                quotation_number=EXCLUDED.quotation_number,
                quotation_amount=EXCLUDED.quotation_amount,
                quotation_currency=EXCLUDED.quotation_currency,
                delivery_days=EXCLUDED.delivery_days,
                quotation_valid_until=EXCLUDED.quotation_valid_until,
                payment_terms=EXCLUDED.payment_terms,
                quotation_notes=EXCLUDED.quotation_notes,
                quotation_status=EXCLUDED.quotation_status,
                quotation_sent_at=COALESCE(EXCLUDED.quotation_sent_at, enquiries.quotation_sent_at),
                payment_due_date=EXCLUDED.payment_due_date,
                payment_email_status=EXCLUDED.payment_email_status,
                payment_receipt_sent_at=COALESCE(EXCLUDED.payment_receipt_sent_at, enquiries.payment_receipt_sent_at),
                payment_reminder_sent_at=COALESCE(EXCLUDED.payment_reminder_sent_at, enquiries.payment_reminder_sent_at),
                updated_at=NOW(), sheet_row=EXCLUDED.sheet_row,
                sync_status='Synced', last_synced_at=NOW()
            RETURNING *
            """,
            {**data, "enquiry_id": enquiry_id, "received_at": received_at, "sheet_row": sheet_row},
        ).fetchone()
        return dict(row)


def get_enquiry(enquiry_id: str) -> dict[str, Any] | None:
    with connection() as conn:
        row = conn.execute("SELECT * FROM enquiries WHERE enquiry_id=%s", (enquiry_id,)).fetchone()
        return dict(row) if row else None
