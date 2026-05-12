# Customer Notes Merger

This project includes:

- a reusable merge engine
- a command-line script
- a modern desktop GUI for Windows packaging now and macOS packaging next

The app takes:

- a semicolon-separated CSV file with customer section header rows
- an Excel file with customer follow-up data in columns `H:J`

It appends these three fields to the matching customer header rows in the CSV:

- `Commentaar`
- `Mail`
- `Whatsapp`

The final output is saved as an Excel file (`.xlsx`).

## What it handles

- Matches by customer name on the section header row.
- Normalizes names to reduce small formatting differences.
- Combines multiple comment, mail, and WhatsApp values for the same customer into one cell per field.
- Preserves all rows from the CSV in the output workbook.

## Files

- [app.py](D:\Razex Solutions LLC (Codex)\PETER\Peter 4\app.py): desktop GUI
- [customer_notes_merger.py](D:\Razex Solutions LLC (Codex)\PETER\Peter 4\customer_notes_merger.py): core merge logic
- [merge_customer_notes.py](D:\Razex Solutions LLC (Codex)\PETER\Peter 4\merge_customer_notes.py): CLI entry point
- [build_windows.ps1](D:\Razex Solutions LLC (Codex)\PETER\Peter 4\build_windows.ps1): Windows build script

## Run The Script

Use the bundled Python runtime or your own Python with `openpyxl` installed:

```powershell
C:\Users\essar\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe `
  .\merge_customer_notes.py `
  ".\Docs\Bestand uit Yuki - Openstaande transacties - one2three.csv" `
  ".\Docs\Openstaande transacties  april (oud).xlsx"
```

The default output file will be created next to the CSV as:

`Bestand uit Yuki - Openstaande transacties - one2three - merged.xlsx`

## Build The Windows App

```powershell
powershell -ExecutionPolicy Bypass -File .\build_windows.ps1
```

This produces:

- [CustomerNotesMerger.exe](D:\Razex Solutions LLC (Codex)\PETER\Peter 4\dist\CustomerNotesMerger.exe)
- [customer_notes_merger.ico](D:\Razex Solutions LLC (Codex)\PETER\Peter 4\assets\customer_notes_merger.ico)

## Build The macOS App

Local macOS build:

```bash
chmod +x ./build_macos.sh
./build_macos.sh
```

GitHub Actions build:

- Workflow file: [.github/workflows/macos-build.yml](D:\Razex Solutions LLC (Codex)\PETER\Peter 4\.github\workflows\macos-build.yml)
- Produces two artifacts:
  - `CustomerNotesMerger-macos-intel`
  - `CustomerNotesMerger-macos-apple-silicon`

### Signing And Notarization

Unsigned macOS apps may trigger Gatekeeper warnings. The workflow is prepared to sign and notarize the app if these GitHub repository secrets are configured:

- `APPLE_CERTIFICATE_P12_BASE64`
- `APPLE_CERTIFICATE_PASSWORD`
- `APPLE_KEYCHAIN_PASSWORD`
- `APPLE_SIGNING_IDENTITY`
- `APPLE_ID`
- `APPLE_TEAM_ID`
- `APPLE_APP_SPECIFIC_PASSWORD`

If those secrets are not present, the workflow still builds the app, but the artifact will be unsigned and may appear suspicious to macOS users.
