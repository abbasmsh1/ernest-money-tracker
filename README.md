# Time & Tune — Ernest Money Tracker

Desktop application (Python / Tkinter) for tracking earnest money records,
generating letters and record PDFs, and sending automatic email reminders
when a deposit becomes due for return.

Main file: `Time_and_Tune_Ernest_Money_Tracker_V8_AUTO_EMAIL_RELIABILITY_FIXED.py`

## Features

- Active and returned earnest money records stored in Excel workbooks
  (`records.xlsx`, `returned_records.xlsx`).
- Letter PDF generated on top of the blank TIME & TUNE template
  (`blank temp 1.pdf`), plus a record summary PDF per customer.
- Automatic reminders: a record is due 15 calendar months after its date.
  The app checks every 5 minutes (configurable via `auto_check_minutes` in
  `settings.json`) and emails the letter and record PDFs to the configured
  recipient, once per reference. Emails are sent on a background thread, so
  the app stays responsive while sending.
- Multi-user: point every PC at the same shared/network data folder.
  Writes are serialized with a cross-process lock file, and email sends are
  claimed per reference so two PCs never send the same reminder twice.
- User accounts with Admin / Employee / Viewer roles
  (PBKDF2-hashed passwords), audit log, and email history log.
- Timestamped backups of all shared data (latest 30 kept), plus an automatic
  backup once per day on the first login.
- Dashboard health strip (email configured, template found, shared folder
  reachable, last auto-check time) and a "Due in 30 days" card; the Due
  Orders window also lists records approaching the 15-month mark.
- Unicode PDFs: bundled DejaVu fonts (`fonts/` folder), so non-Latin
  customer names print correctly in letters and records.

## Requirements

- Python 3.9+ with Tkinter
- Packages: `pandas`, `openpyxl`, `fpdf2`, `python-dateutil`,
  `reportlab`, `pypdf`

```
python3 -m venv .venv
.venv/bin/pip install pandas openpyxl fpdf2 python-dateutil reportlab pypdf
```

(On systems that allow it, a plain
`pip install pandas openpyxl fpdf2 python-dateutil reportlab pypdf` works too.)

## Running

```
.venv/bin/python Time_and_Tune_Ernest_Money_Tracker_V8_AUTO_EMAIL_RELIABILITY_FIXED.py
```

### Windows exe (no Python needed on office PCs)

On a Windows PC with Python, run `build_windows_exe.bat` in this folder.
It produces `dist\ErnestMoneyTracker.exe` with the letter template and fonts
bundled inside; copy that single file to any PC.

First run creates the first Admin account and a local data folder
(`~/TimeAndTune_Ernest_Data`). Use the shared-folder button on the login
screen to point the PC at a shared/network folder for multi-user setups.

Keep the letter template PDF next to the script, or select it once when
prompted; the chosen path is remembered in the data folder
(`template_path.txt`).

## Email configuration

Configure via the Email Settings screen (Admin) or environment variables:

| Variable | Purpose |
|----------|---------|
| `ERNEST_SMTP_HOST` | SMTP host (default `smtp.gmail.com`) |
| `ERNEST_SMTP_PORT` | SMTP port (default `465`) |
| `ERNEST_SMTP_SECURITY` | `SSL` or `STARTTLS` |
| `ERNEST_SENDER_EMAIL` | Sender account |
| `ERNEST_RECIPIENT_EMAIL` | Reminder recipient |
| `ERNEST_APP_PASSWORD` | Sender app password (overrides saved one) |

For Gmail, use a Google App Password, not the normal account password.

The password entered in Email Settings is stored per machine in
`~/TimeAndTune_Ernest_Client/smtp_password.txt` (owner-only permissions),
never in the shared data folder. On a multi-PC setup, enter it once on each
machine that should send reminders, or set `ERNEST_APP_PASSWORD` instead.

## Data layout

| Path | Contents |
|------|----------|
| shared data folder | `records.xlsx`, `returned_records.xlsx`, `users.xlsx`, `audit_log.xlsx`, `email_history.xlsx`, `settings.json`, `Records/<customer>/` PDFs, `Backups/` |
| `~/TimeAndTune_Ernest_Client/` | per-machine config: shared folder path, SMTP password |
