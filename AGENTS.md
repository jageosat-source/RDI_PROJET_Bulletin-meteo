# Repository Guidelines

## Project Structure & Module Organization

This repository contains the production weather bulletin pipeline for Bordeaux.
The active production version is `20260623`; source code lives in
`20260623/Programme`.

- `20260623/Programme/generate_bulletin.py`: fetches Open-Meteo data and writes the HTML bulletin.
- `20260623/Programme/html_to_pdf.py`: converts the latest or selected HTML bulletin to PDF with Playwright/Chromium.
- `20260623/Programme/send_bulletin.ps1`: sends the PDF by Outlook e-mail.
- `20260623/Programme/run_bulletin_complet.ps1`: single production entry point.
- `20260623/Programme/Liste-diffusion.csv`: editable recipient list, used in BCC.
- `20260623/Bulletins`: runtime HTML/PDF outputs, excluded from Git except `.gitkeep`.
- `20260623/Archives`: local proofs, logs, and release ZIPs, excluded from Git except `.gitkeep`.
- `docs/`: production and archiving notes.

## Build, Test, and Development Commands

Run all commands from PowerShell.

```powershell
& "D:\JAU\RDI_PROJET_Bulletin-meteo\20260623\Programme\run_bulletin_complet.ps1" -DryRun
```
Generates HTML and PDF without sending e-mail. Use before every commit.

```powershell
& "D:\JAU\RDI_PROJET_Bulletin-meteo\20260623\Programme\run_bulletin_complet.ps1"
```
Runs the production pipeline and sends the PDF through Outlook.

```powershell
& "D:\JAU\RDI_PROJET_Bulletin-meteo\20260623\Programme\setup_tache_planifiee.ps1"
```
Creates or updates the weekday 08:00 scheduled task.

```powershell
& "D:\JAU\.venv\Scripts\python.exe" -m py_compile `
  "D:\JAU\RDI_PROJET_Bulletin-meteo\20260623\Programme\generate_bulletin.py" `
  "D:\JAU\RDI_PROJET_Bulletin-meteo\20260623\Programme\html_to_pdf.py"
```
Checks Python syntax with the portable runtime used in production.

## Coding Style & Naming Conventions

Use ASCII for new code unless the target file already requires UTF-8 text.
Python uses 4-space indentation, `pathlib`, `argparse`, explicit errors, and small focused functions.
PowerShell scripts should keep `$ErrorActionPreference = "Stop"`, clear step logging, and explicit paths.
Keep production filenames stable; generated bulletins follow `Meteo_YYYYMMDD-HHMM.html` and `.pdf`.

## Testing Guidelines

There is no separate automated test suite. Validation is operational:
run `-DryRun`, confirm HTML/PDF creation in `20260623/Bulletins`, and inspect `run_log.txt`.
For mail changes, perform one controlled real send and archive proof under
`20260623/Archives/Preuves/YYYYMMDD`.

## Commit & Pull Request Guidelines

Use short imperative commit messages, matching history examples such as
`Add Outlook mail timeout safeguard` or `Finalize RDI MeteoBot production release`.
Production tags follow `RDI_MeteoBot_YYYYMMDD-vX.Y.Z`.

Pull requests or release notes must include the change summary, validation command used,
mail/schedule impact, and any archive or version tag created.

## Security & Configuration Tips

Do not commit generated bulletins, logs, release ZIPs, or proof archives.
Keep recipient changes in `Liste-diffusion.csv`; e-mail recipients must remain in BCC.
The Teams channel and HTML e-mail attachment are intentionally abandoned.
Do not rely on the Windows `PATH` for Python; use `D:\JAU\.venv\Scripts\python.exe`.
