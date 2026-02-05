# Venmo Email Parser

Parses Venmo payment notification emails from Gmail and logs them to an Excel spreadsheet.

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure Gmail API

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project (or select existing)
3. Enable the Gmail API
4. Create OAuth 2.0 credentials (Desktop application)
5. Download the credentials and save as `credentials.json` in this directory

### 3. Create Gmail label

Create a label called "Venmo" in Gmail and apply it to your Venmo payment emails.

### 4. Excel file

Ensure `Poker.xlsx` exists with a sheet named `money ` (note the trailing space). The script writes to columns B (name) and C (amount).

## Usage

### Single run

```bash
python run_etl.py --once
```

Processes all new emails once and exits.

### Continuous polling

```bash
python run_etl.py
```

Runs continuously, checking for new emails every 5 minutes. Press Ctrl+C to stop.

## How it works

- Authenticates with Gmail using OAuth 2.0
- Fetches emails from the "Venmo" label
- Parses payment info from email body:
  - Incoming: "Name paid you $X" (positive amount)
  - Outgoing: "You paid Name $X" (negative amount)
- Appends payments to Excel spreadsheet
- Tracks processed message IDs in `processed_messages.json` to avoid duplicates

## Files

| File | Purpose |
|------|---------|
| `run_etl.py` | Main script |
| `credentials.json` | Gmail API credentials (you provide) |
| `token.pickle` | Cached auth token (auto-generated) |
| `processed_messages.json` | Tracks processed emails (auto-generated) |
| `Poker.xlsx` | Output spreadsheet (you provide) |
