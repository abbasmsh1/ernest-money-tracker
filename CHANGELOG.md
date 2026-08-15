# Changelog

## 2026-08-15 — New features

### Added
- Automatic daily backup: the first login of each day creates a timestamped
  backup of all shared data (runs in the background, audited as AUTO BACKUP).
- Dashboard health strip: shared folder reachable, email configured,
  template found, and time of the last automatic reminder check.
- "Due in 30 days": a dashboard card counting records that reach the
  15-month mark within 30 days, and a matching list in the Due Orders
  window below the overdue records.
- Unicode PDF support: DejaVu fonts bundled in `fonts/` are used for both
  the record PDF and the letter, so non-Latin customer names no longer
  print as `?`. Falls back to the old Latin-1 behavior if the fonts folder
  is missing. Requires `fpdf2` (replaces the old `fpdf` package).
- `build_windows_exe.bat`: builds a single-file Windows exe (PyInstaller)
  with the template and fonts bundled, for PCs without Python. The template
  and font lookup now also work inside a frozen exe.

## 2026-08-15 — Responsiveness and reliability

### Fixed
- The UI no longer freezes while reminder emails are sent. Manual reminder
  checks, the automatic 5-minute check, and Email History resends all run on
  a background thread; result dialogs appear when sending finishes. Only one
  reminder batch runs at a time.
- The `auto_check_minutes` setting is now honored. Previously the automatic
  check interval was hardcoded to 5 minutes and the setting was ignored.
- Error dialogs are no longer shown while the shared-data lock is held.
  Previously a user leaving an error dialog open (for example, a workbook
  open in Excel) could block every other PC for the lock timeout. Email and
  audit writes now log errors to the console / `email_errors.log` instead of
  popping dialogs; interactive saves show the dialog after the lock is
  released.
- The template file picker can no longer appear from a background thread.
  If the template is missing during an automatic check, the send fails and
  is recorded in Email History instead.

### Removed
- The vestigial `emailed_customers.txt` write path (`mark_emailed`).
  Duplicate-send prevention has used the email history workbook since V8;
  the text file was written but never read.

## 2026-08-15 — Cleanup and security fix

### Fixed
- Letter generation now uses the template resolver (`resolve_template_path()`)
  instead of a hardcoded path next to the script. The resolver honors a
  previously saved template path, tolerates URL-encoded filenames, searches
  common folders, and falls back to a one-time file picker whose choice is
  remembered in `template_path.txt`. Previously this logic existed but was
  never called.

### Security
- SMTP app password is no longer stored in plaintext in the shared data
  folder, where every network user could read it. It is now stored per
  machine in `~/TimeAndTune_Ernest_Client/smtp_password.txt` with owner-only
  permissions (0600). On first launch after this change, an existing shared
  password file is migrated to the local machine and the shared copy is
  deleted; other PCs must re-enter the password once in Email Settings or
  set `ERNEST_APP_PASSWORD`.
- Backups no longer include the SMTP password file. Older backup folders
  (`Backups/*/smtp_password.txt`) may still contain the old password; delete
  them or rotate the app password.

### Removed
- Dead code left over from V7 (~70 lines): the old environment-variable-only
  `send_email()` and `email_configured()` (referenced undefined globals and
  would have crashed if called), the old customer-based `mark_emailed()`,
  and the unused `load_emailed_customers()`. All were shadowed or unreachable;
  behavior is unchanged.

## V8 — Auto email reliability

- Reference-based reminder tracking, per-reference email send claims to
  prevent duplicate sends across PCs, email history log with attempt counts,
  and retry/finalize logic around SMTP sends.
- Multi-user shared data folder with cross-process lock file, user accounts
  with roles, audit log, and automatic timestamped backups.

(History before V8 is not recorded.)
