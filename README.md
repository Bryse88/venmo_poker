# Venmo Email Parser

Parses Venmo payment notification emails from Gmail and logs them to a CSV file.

## Features

- Parses both incoming ("X paid you") and outgoing ("You paid X") payments
- Captures payment date and note/memo
- Handles Gmail pagination (processes all emails, not just first 100)
- Tracks processed emails to avoid duplicates
- Configurable via environment variables
- Proper logging with configurable levels

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

### 4. Configuration (optional)

Copy `.env.example` to `.env` and customize:

```bash
cp .env.example .env
```

Available settings:

| Variable | Default | Description |
|----------|---------|-------------|
| `CSV_PATH` | `payments.csv` | Path to output CSV file |
| `GMAIL_LABEL` | `Venmo` | Gmail label to search |
| `POLL_INTERVAL` | `300` | Seconds between checks (continuous mode) |
| `LOG_LEVEL` | `INFO` | Logging level (DEBUG, INFO, WARNING, ERROR) |

## Output

The CSV file has these columns:

| Column | Data |
|--------|------|
| Name | Person's name |
| Amount IN | Incoming payment amount |
| Amount OUT | Outgoing payment amount |
| Date | Payment date/time |
| Note | Payment memo |

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

Runs continuously, checking for new emails every 5 minutes (configurable). Press Ctrl+C to stop.

### Debug mode

```bash
LOG_LEVEL=DEBUG python run_etl.py --once
```

## Files

| File | Purpose |
|------|---------|
| `run_etl.py` | Main script |
| `requirements.txt` | Python dependencies |
| `.env` | Configuration (you create from `.env.example`) |
| `credentials.json` | Gmail API credentials (you provide) |
| `token.pickle` | Cached auth token (auto-generated) |
| `processed_messages.json` | Tracks processed emails (auto-generated) |
| `payments.csv` | Output file (auto-generated) |
