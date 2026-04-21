# Release Notes

## 1.0.1

### Both platforms

- **Per-invocation LibreOffice profile** — each conversion now runs against its own isolated user profile directory (`-env:UserInstallation=...`). This fixes the profile-lock contention that caused most files to fail silently when converting in parallel.
- **Profile template prewarm** — on the first convert after app start, the app initializes a template profile once (with `--terminate_after_init`) and then copies it for each subsequent conversion. Avoids paying the first-run setup cost per file.
- **Concurrency cap at 8** — soffice doesn't scale past this, so capping reduces contention without reducing throughput.
- **Default Paper Size option** (Letter / A4 / Legal) under Options → Default Paper Size. Injected into each profile via `registrymodifications.xcu` so LibreOffice uses this instead of querying a system printer for page defaults. Also sets `PrinterIndependentLayout=2` for deterministic output.
- **Extra soffice startup flags** — `--norestore --nologo --nofirststartwizard --nodefault --nolockcheck` for clean, non-interactive headless behavior.

### Windows

- **Debug Logging toggle** under Options → Debug Logging. When enabled, logs per-process detail (PID, elapsed time, exit code, semaphore state, profile setup) for diagnosing conversion issues.
- **Tracked process cleanup** — the app now tracks only the soffice processes *it* launched and cleans them up on convert start and form close. Removes the risk of killing unrelated LibreOffice sessions.
- Settings file format extended to `key=value` lines (backwards-compatible read of the old single-line theme file).

### Linux

- Settings persisted to `~/.config/LibreDocToPdf/settings.txt`.

### Bug fixes

- LibreOffice conversions no longer silently fail when multiple files are processed in parallel (the profile-lock issue above).
- Orphaned soffice processes from previous runs no longer accumulate on Windows.
