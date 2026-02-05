#!/usr/bin/env python3
"""
Venmo Email Parser for Gmail
Reads Venmo payment notification emails and adds them to Excel spreadsheet
"""

from __future__ import annotations

import argparse
import base64
import csv
import json
import logging
import os
import pickle
import re
import time
from dataclasses import dataclass
from datetime import datetime
from email.utils import parsedate_to_datetime
from typing import Optional

from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build, Resource

# Load environment variables
load_dotenv()

# Configuration
CSV_PATH = os.getenv('CSV_PATH', 'payments.csv')
GMAIL_LABEL = os.getenv('GMAIL_LABEL', 'Venmo')
POLL_INTERVAL = int(os.getenv('POLL_INTERVAL', '300'))
LOG_LEVEL = os.getenv('LOG_LEVEL', 'INFO')

# Constants
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
PROCESSED_IDS_FILE = 'processed_messages.json'

# Setup logging
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL.upper()),
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


@dataclass
class Payment:
    """Represents a Venmo payment."""
    name: str
    amount: float
    date: Optional[datetime] = None
    note: Optional[str] = None

    @property
    def is_outgoing(self) -> bool:
        return self.amount < 0


def load_processed_ids() -> set[str]:
    """Load set of already processed message IDs from JSON file."""
    if os.path.exists(PROCESSED_IDS_FILE):
        with open(PROCESSED_IDS_FILE, 'r') as f:
            return set(json.load(f))
    return set()


def save_processed_ids(processed_ids: set[str]) -> None:
    """Save processed message IDs to JSON file."""
    with open(PROCESSED_IDS_FILE, 'w') as f:
        json.dump(list(processed_ids), f)


def authenticate_gmail() -> Resource:
    """Authenticate with Gmail API and return service object."""
    creds: Optional[Credentials] = None

    # Token file stores user's access and refresh tokens
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If no valid credentials, let user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save credentials for next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('gmail', 'v1', credentials=creds)


def get_venmo_emails(service: Resource, label_name: str = GMAIL_LABEL) -> list[dict]:
    """Fetch all emails from Venmo label, handling pagination."""
    try:
        # Get label ID for specified label
        labels = service.users().labels().list(userId='me').execute()
        label_id: Optional[str] = None

        for label in labels.get('labels', []):
            if label['name'].lower() == label_name.lower():
                label_id = label['id']
                break

        if not label_id:
            logger.warning(f"Label '{label_name}' not found in Gmail")
            return []

        # Fetch all messages with pagination
        all_messages: list[dict] = []
        page_token: Optional[str] = None

        while True:
            results = service.users().messages().list(
                userId='me',
                labelIds=[label_id],
                q='from:venmo@venmo.com',
                pageToken=page_token
            ).execute()

            messages = results.get('messages', [])
            all_messages.extend(messages)

            page_token = results.get('nextPageToken')
            if not page_token:
                break

            logger.debug(f"Fetched {len(all_messages)} messages, getting next page...")

        logger.info(f"Found {len(all_messages)} total email(s) in '{label_name}' label")
        return all_messages

    except Exception as e:
        logger.error(f"Error fetching emails: {e}")
        return []


def extract_payment_info(email_body: str) -> tuple[Optional[str], Optional[float], Optional[str]]:
    """Extract payer name, amount, and note from Venmo email.

    Returns (name, amount, note) where:
    - Incoming payments (X paid you): positive amount
    - Outgoing payments (You paid X): negative amount
    - Note: the payment memo if found
    """
    name: Optional[str] = None
    amount: Optional[float] = None
    note: Optional[str] = None

    # Pattern for "FirstName LastName paid you $XX.XX" (incoming)
    incoming_pattern = r'([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)\s+paid you\s+\$(\d+(?:\.\d{2})?)'

    # Pattern for "You paid FirstName LastName $XX.XX" (outgoing)
    outgoing_pattern = r'You paid\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)\s+\$(\d+(?:\.\d{2})?)'

    # Pattern for payment note - typically in quotes or after a dash
    note_patterns = [
        r'["\u201c]([^"\u201d]+)["\u201d]',  # "note" or "note"
        r'[-\u2013\u2014]\s*(.+?)(?:\n|$)',   # - note or — note
    ]

    # Check for incoming payment
    match = re.search(incoming_pattern, email_body)
    if match:
        name = match.group(1)
        amount = float(match.group(2))
    else:
        # Check for outgoing payment
        match = re.search(outgoing_pattern, email_body)
        if match:
            name = match.group(1)
            amount = -float(match.group(2))  # Negative for outgoing

    # Try to extract note
    for pattern in note_patterns:
        note_match = re.search(pattern, email_body)
        if note_match:
            potential_note = note_match.group(1).strip()
            # Filter out common non-note matches
            if potential_note and len(potential_note) > 1 and not potential_note.startswith('http'):
                note = potential_note
                break

    return name, amount, note


def get_email_date(message: dict) -> Optional[datetime]:
    """Extract date from email message headers."""
    headers = message.get('payload', {}).get('headers', [])

    for header in headers:
        if header['name'].lower() == 'date':
            try:
                return parsedate_to_datetime(header['value'])
            except Exception:
                pass

    # Fallback to internalDate (milliseconds since epoch)
    internal_date = message.get('internalDate')
    if internal_date:
        try:
            return datetime.fromtimestamp(int(internal_date) / 1000)
        except Exception:
            pass

    return None


def parse_email_content(service: Resource, message_id: str) -> Optional[Payment]:
    """Get email content and extract payment info."""
    try:
        message = service.users().messages().get(
            userId='me',
            id=message_id,
            format='full'
        ).execute()

        # Extract email body
        parts = message.get('payload', {}).get('parts', [])
        email_body = ''

        # Handle different email structures
        if parts:
            for part in parts:
                if part.get('mimeType') == 'text/plain':
                    data = part.get('body', {}).get('data', '')
                    if data:
                        email_body = base64.urlsafe_b64decode(data).decode('utf-8')
                        break
        else:
            # Email body might be directly in payload
            data = message.get('payload', {}).get('body', {}).get('data', '')
            if data:
                email_body = base64.urlsafe_b64decode(data).decode('utf-8')

        # Also check snippet for simpler emails
        if not email_body:
            email_body = message.get('snippet', '')

        name, amount, note = extract_payment_info(email_body)

        if name and amount is not None:
            return Payment(
                name=name,
                amount=amount,
                date=get_email_date(message),
                note=note
            )

        return None

    except Exception as e:
        logger.error(f"Error parsing email {message_id}: {e}")
        return None


def add_to_csv(csv_path: str, payments: list[Payment]) -> bool:
    """Add payment data to CSV file. Returns True on success."""
    try:
        # Check if file exists to determine if we need headers
        file_exists = os.path.exists(csv_path)

        with open(csv_path, 'a', newline='') as f:
            writer = csv.writer(f)

            # Write header if new file
            if not file_exists:
                writer.writerow(['Name', 'Amount IN', 'Amount OUT', 'Date', 'Note'])

            # Write each payment
            for payment in payments:
                amount_in = payment.amount if not payment.is_outgoing else ''
                amount_out = abs(payment.amount) if payment.is_outgoing else ''
                date_str = payment.date.strftime('%Y-%m-%d %H:%M:%S') if payment.date else ''
                writer.writerow([payment.name, amount_in, amount_out, date_str, payment.note or ''])

        logger.info(f"Added {len(payments)} payment(s) to CSV")
        return True

    except Exception as e:
        logger.error(f"Error updating CSV: {e}")
        return False


def run_parser(service: Resource, csv_path: str, processed_ids: set[str]) -> int:
    """Run a single parsing cycle. Returns number of new payments processed."""
    messages = get_venmo_emails(service)

    if not messages:
        logger.info("No Venmo emails found")
        return 0

    # Filter out already processed messages
    new_messages = [msg for msg in messages if msg['id'] not in processed_ids]

    if not new_messages:
        logger.info("No new emails to process")
        return 0

    logger.info(f"Processing {len(new_messages)} new email(s)")

    payments: list[Payment] = []
    for msg in new_messages:
        payment = parse_email_content(service, msg['id'])
        if payment:
            payments.append(payment)
            direction = "paid you" if not payment.is_outgoing else "you paid"
            amount_display = abs(payment.amount)
            note_display = f' - "{payment.note}"' if payment.note else ''
            logger.info(f"  Found: {payment.name} {direction} ${amount_display:.2f}{note_display}")
        else:
            logger.warning(f"  Could not parse message {msg['id']}")

        # Mark as processed regardless of parse success
        processed_ids.add(msg['id'])

    if payments:
        add_to_csv(csv_path, payments)
    else:
        logger.warning("No payment information could be extracted from new emails")

    return len(payments)


def main() -> None:
    """Main entry point."""
    parser = argparse.ArgumentParser(description='Venmo Email Parser for Gmail')
    parser.add_argument('--once', action='store_true',
                        help='Run once and exit instead of continuous polling')
    args = parser.parse_args()

    logger.info("Venmo Email Parser starting")
    logger.info(f"Config: csv={CSV_PATH}, label={GMAIL_LABEL}, poll={POLL_INTERVAL}s")

    logger.info("Authenticating with Gmail...")
    service = authenticate_gmail()
    logger.info("Authentication successful")

    # Load previously processed message IDs
    processed_ids = load_processed_ids()
    logger.info(f"Loaded {len(processed_ids)} previously processed message ID(s)")

    if args.once:
        # Single run mode
        run_parser(service, CSV_PATH, processed_ids)
        save_processed_ids(processed_ids)
        logger.info("Complete!")
    else:
        # Continuous polling mode
        logger.info(f"Starting continuous polling (every {POLL_INTERVAL}s, Ctrl+C to stop)")
        try:
            while True:
                run_parser(service, CSV_PATH, processed_ids)
                save_processed_ids(processed_ids)
                logger.debug(f"Sleeping for {POLL_INTERVAL} seconds...")
                time.sleep(POLL_INTERVAL)
        except KeyboardInterrupt:
            logger.info("Stopping parser...")
            save_processed_ids(processed_ids)
            logger.info("Processed IDs saved. Goodbye!")


if __name__ == '__main__':
    main()
