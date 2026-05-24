# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

**Kiosco Seguimientos** is a touchscreen kiosk for physiotherapy clinics. Physiotherapists enter their registration number on a tablet to register first-visit follow-ups. A daily trigger sends email reminders when a re-evaluation is due (every 10 days), with one-click actions to renew or close the follow-up.

## Stack

- **Frontend**: Single-file HTML/CSS/JS — no framework, no bundler, no npm
- **Backend**: Google Apps Script (deployed as a Web App)
- **Storage**: Google Sheets (two sheets: `Fisios` and `Seguimientos`)
- **Notifications**: Gmail via `MailApp.sendEmail`
- **Clasp**: local ↔ GAS sync via CLI

## Structure

```
Codigo.gs        ← Backend (Apps Script)
index.html       ← Kiosk UI (embedded CSS + JS)
appsscript.json  ← GAS project configuration
setup.md         ← Human-facing deployment guide
```

## Running / deploying

No build step. Use `make` to interact with GAS:

```bash
make push-personal                         # push to personal/dev GAS project
make push-prod                             # push to production GAS project (restores personal after)
make pull                                  # pull from personal project
make deploy ID=<deploymentId> DESC="v1.x" # push + publish to production URL (restores personal after)
make open                                  # open personal GAS editor
make logs                                  # view personal execution logs
```

`.clasp.personal.json` is the default environment — `push-prod` and `deploy` restore it after running so `.clasp.json` always points to personal between commands.

See [setup.md](setup.md) for the full deployment guide including how to create the Spreadsheet and install the daily trigger.

## Key backend functions (`Codigo.gs`)

| Function | Purpose |
|---|---|
| `doGet(e)` | Entry point — routes email actions or serves the kiosk HTML |
| `validarColegiado(num)` | Looks up a physio by registration number in the `Fisios` sheet |
| `registrarSeguimiento(num, fecha, hora)` | Appends a new row to `Seguimientos`; prevents duplicates |
| `obtenerSeguimientosFisio(num)` | Returns active follow-ups for a given physio |
| `enviarRecordatoriosDiarios()` | Daily trigger — sends reminder emails when re-evaluation is due |
| `cerrarSeguimientoDesdeEmail(fila)` | Sets status to `'finalizado'` from email link |
| `reevaluarDesdeEmail(fila)` | Advances `Próxima_Reev` by +10 days from email link |
| `instalarTrigger()` | Run once from the editor to install the daily 8:00 trigger |

## Constants to configure

```js
const DIAS_SEGUIM = 10;  // days between reminders (line 9 of Codigo.gs)
```

`SHEET_ID` is **not in the code** — it is stored as a GAS Script Property:
**Project settings → Script properties → Key: `SHEET_ID`**

## Sheets schema

**Fisios:** `A: Nombre | B: Num_Colegiado | C: Email`

**Seguimientos:** `A: ID | B: Fisio | C: Num_Colegiado | D: Email | E: Fecha_Visita | F: Hora_Visita | G: Estado | H: Fecha_Registro | I: Próxima_Reev`

## Clasp files

- `.clasp.prod.json` — versioned; contains the production Script ID (no credentials)
- `.clasp.json` — active file used by clasp; **not versioned** (in `.gitignore`); copy from `.clasp.prod.json`
- `.clasprc.json` — Google OAuth token; **never commit** (lives in `~/.clasprc.json`)
- `appsscript.json` — GAS project manifest (timezone, runtime, permissions); versioned

## Important development note

`google.script.run` does not exist outside GAS. You can edit HTML/CSS locally, but to test any backend call you must `clasp push` and open the `/dev` URL.

## Commit format

```
git commit -m "short imperative title" -m "description when necessary"
```

- The first `-m` is the title (max ~72 characters)
- The second `-m` is only included when there is relevant context to add
- Never use `git commit` without flags or interactive editors
- Never add co-authorship (`Co-Authored-By`) under any circumstances
