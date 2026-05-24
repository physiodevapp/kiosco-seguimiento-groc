# Kiosco Seguimientos · Setup Guide

## Project files

```
kiosco-seguimientos/
├── Codigo.gs        ← Backend (Apps Script)
├── index.html       ← Touchscreen kiosk
├── appsscript.json  ← GAS project configuration
├── README.md
└── setup.md
```

---

## 1. Google Sheets — Create the Spreadsheet

Create a new Google Sheets file and note the ID from the URL:
`https://docs.google.com/spreadsheets/d/**ID_HERE**/edit`

### Sheet "Fisios"
Create a sheet with this exact name and these columns:

| A: Nombre | B: Num_Colegiado | C: Email |
|---|---|---|
| Ana García | 28001 | ana@clinica.com |
| Luis Pérez | 28002 | luis@clinica.com |

> Only the administrator adds physiotherapists here. This is not done from the kiosk.

### Sheet "Seguimientos"
Create a sheet with this exact name. Columns are filled in automatically:

| A: ID | B: Fisio | C: Num_Colegiado | D: Email | E: Fecha_Visita | F: Hora_Visita | G: Estado | H: Fecha_Registro | I: Próxima_Reev |
|---|---|---|---|---|---|---|---|---|

---

## 2. Apps Script — Configure the project

1. In the Spreadsheet, go to **Extensions → Apps Script**
2. Clear the contents of `Código.gs` and paste the contents of `Codigo.gs`
3. Create a new HTML file: **File → New → HTML**, name it `index`
4. Paste the contents of `index.html`

### Set the Spreadsheet ID

The `SHEET_ID` is **not stored in the code**. Set it as a Script Property in GAS:

1. In the Apps Script editor: **Project settings → Script properties → Add property**
2. Key: `SHEET_ID`
3. Value: your Spreadsheet ID (the long string from the Sheets URL)

This keeps the ID out of the repository.

---

## 3. Deploy as a Web App

1. In Apps Script: **Deploy → New deployment**
2. Type: **Web app**
3. Execute as: **Me** (your Google account)
4. Who has access: **Anyone** (so the tablet works without login)
5. Click **Deploy** → copy the generated URL

> This URL is what you open in the tablet browser in kiosk mode.

---

## 4. Install the daily trigger (once only)

1. In the Apps Script editor, select the `instalarTrigger` function
2. Click ▶ **Run**
3. Accept the requested permissions (email, Sheets)
4. The trigger will be installed and will send reminders every day at 8:00

You can verify it under **Triggers** (clock icon in the side menu).

---

## 5. Set up the tablet in kiosk mode

### Android (recommended)
- Enable **screen pinning** (Settings → Security)
- Or use a kiosk app such as **Kiosk Browser** or **Fully Kiosk Browser**
- Point the browser to the Web App URL

### iPad
- Use **Guided Access** (Settings → Accessibility → Guided Access)
- Open Safari with the Web App URL

---

## 6. GitHub — Version control

### Workflow

Since Apps Script does not sync natively with GitHub, use **clasp**:

```bash
# Install clasp
npm install -g @google/clasp

# Login
clasp login

# Clone the existing project (using the Script ID from Apps Script)
clasp clone <SCRIPT_ID>

# Push local changes to Apps Script
clasp push

# Pull changes from Apps Script to local
clasp pull
```

The Script ID is found in Apps Script: **Project → Project settings → Script ID**.

### Recommended .gitignore
```
.clasp.json
```

> `.clasp.json` contains the Script ID — no need to version it.

---

## 7. Full system flow

```
Tablet (kiosk)
└── Physio enters registration number
    └── Selects date and time of first visit
        └── Recorded in Google Sheets

Daily trigger (8:00)
└── Checks active follow-ups
    └── If today >= Próxima_Reev → sends email to physio
        └── Email with two buttons:
            ├── ✅ Re-evaluated → adds +10 days in the Sheet
            └── 🔴 Close → sets status to "finalizado"
```

---

## 8. Customisation

| Variable | File | Description |
|---|---|---|
| `DIAS_SEGUIM` | Codigo.gs line 9 | Days between reminders (default: 10) |
| `atHour(8)` | Codigo.gs function `instalarTrigger` | Time the daily email is sent |
| Colours and typography | index.html (CSS variables) | Visual palette of the kiosk |

---

## Support

If the trigger is not sending emails:
1. Check that the physio has an email address in the "Fisios" sheet
2. Check Apps Script **View → Execution log**
3. Make sure you accepted the Gmail permissions when running `instalarTrigger`
