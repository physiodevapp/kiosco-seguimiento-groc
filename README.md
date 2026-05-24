# Kiosco Seguimientos · Physiotherapy

A touchscreen kiosk for registering first visits and managing automated patient re-evaluation reminders.

## Stack

- **Frontend**: HTML/CSS/JS — touchscreen kiosk UI (tablet)
- **Backend**: Google Apps Script + Google Sheets
- **Notifications**: Gmail (daily trigger)
- **Clasp**: local ↔ GAS sync via CLI

## Structure

```
Codigo.gs        ← Backend (Apps Script)
index.html       ← Kiosk UI
appsscript.json  ← GAS project configuration
setup.md         ← Full setup guide
```

## Setup

See [setup.md](setup.md) for the complete deployment guide.

## Workflow

```bash
# Edit locally, then:
clasp push                                            # → Google Apps Script
git add . && git commit -m "..." && git push         # → GitHub

# Publish to production
clasp deploy --deploymentId <id> --description "vX"
```

## Flow

1. Physiotherapist enters their registration number on the tablet
2. Selects date and time of first visit → recorded in Google Sheets
3. A daily trigger runs at 8:00 and checks active follow-ups
4. When re-evaluation is due (every 10 days), the physio receives an email with two options:
   - **Re-evaluated** → adds +10 days to the next review date
   - **Close** → marks the follow-up as finished
