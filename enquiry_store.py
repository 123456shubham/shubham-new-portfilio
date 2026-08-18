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
                last_synced_at TIMESTAMPTZ
            )
            """
        )
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS amount_received NUMERIC(14, 2) NOT NULL DEFAULT 0")
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS payment_date DATE")
        conn.execute("ALTER TABLE enquiries ADD COLUMN IF NOT EXISTS payment_status TEXT NOT NULL DEFAULT 'Pending'")
        conn.execute(
            "CREATE INDEX IF NOT EXISTS enquiries_updated_at_idx ON enquiries(updated_at DESC)"
        )


def parse_amount(value: Any) -> Decimal | None:
    text = str(value or "").replace(",", "").replace("₹", "").strip()
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
                lead_status, email_status, source, validity, validation_notes
            ) VALUES (
                %(enquiry_id)s, %(submission_token)s, %(received_at)s, %(client_name)s,
                %(email)s, %(phone)s, %(project_subject)s, %(project_details)s,
                %(project_amount)s, 'New', 'Pending', 'Portfolio Website',
                %(validity)s, %(validation_notes)s
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
            ) VALUES (
                %(enquiry_id)s, %(received_at)s, %(client_name)s, %(email)s, %(phone)s,
                %(project_subject)s, %(project_details)s, %(project_amount)s,
                %(amount_received)s, %(payment_date)s, %(payment_status)s,
                %(lead_status)s, %(email_status)s, %(source)s, %(validity)s,
                %(validation_notes)s, NOW(), %(sheet_row)s, 'Synced', NOW()
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
