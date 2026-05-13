# CUSTOMER NOTES MERGER

This package contains the desktop application builds for Windows and macOS, along with sample data files.

## Included Files

- `CustomerNotesMerger-windows\CustomerNotesMerger.exe`
  Windows version

- `CustomerNotesMerger-macos-intel-dmg\Customer Notes Merger-intel.dmg`
  Mac version for Intel-based Macs

- `CustomerNotesMerger-macos-apple-silicon-dmg\Customer Notes Merger-apple-silicon.dmg`
  Mac version for Apple Silicon Macs such as `M1`, `M2`, `M3`, or `M4`

- `Data\Original`
  Contains the original sample CSV and Excel source files

- `Data\Processed`
  Contains the processed sample output file

## What This App Does

The app takes:

- one CSV file
- one Excel file

It matches customers by the customer section header row, copies Excel columns `H` to `J`, and creates a new Excel output file.

The output includes:

- copied `Commentaar`, `Mail`, and `Whatsapp` values
- combined repeated comment lines into one field when needed
- wrapped text in the comments column for easier reading

The original files are not changed.

## Which Version To Use

### Windows

Use:

- `CustomerNotesMerger-windows\CustomerNotesMerger.exe`

### Mac Intel

Use:

- `CustomerNotesMerger-macos-intel-dmg\Customer Notes Merger-intel.dmg`

This is for older Macs with Intel processors.

### Mac Apple Silicon

Use:

- `CustomerNotesMerger-macos-apple-silicon-dmg\Customer Notes Merger-apple-silicon.dmg`

This is for newer Macs with Apple chips such as `M1`, `M2`, `M3`, or `M4`.

## How To Check Mac Type

1. Click the Apple menu.
2. Click `About This Mac`.
3. Check the processor or chip.

- If it says `Intel`, use the Intel DMG.
- If it says `Apple M1`, `M2`, `M3`, or `M4`, use the Apple Silicon DMG.

## How To Install

### Windows

1. Open `CustomerNotesMerger-windows`.
2. Double-click `CustomerNotesMerger.exe`.
3. If Windows shows a security prompt, choose `More info` and then `Run anyway` if you trust the file source.

### Mac

1. Open the correct DMG for your Mac type.
2. Drag the app to `Applications` if macOS shows that option.
3. Open the app from `Applications`.
4. If Gatekeeper blocks the app, right-click the app, choose `Open`, and confirm.

## How To Use

1. Open the application.
2. Click `Browse CSV` and select the CSV file.
3. Click `Browse Excel` and select the Excel file.
4. Click `Save As` and choose where the new Excel output should be saved.
5. Click `Create Merged Workbook`.

## What Happens After Running

The app will:

- read the CSV file
- read the Excel file
- match customers by the customer header row
- copy Excel columns `H` to `J`
- create a new Excel workbook as output

## Sample Files

You can review the included sample files here:

- `Data\Original`
- `Data\Processed`

************************************************************
***** RAZEX SOLUTIONS LLC *****
Website : ai.razexsolutions.com
Contact : +1 307 220 4428
Email   : info@razexsolutions.com

We build AI automation, custom software, internal tools,
data processing systems, workflow automation, and desktop apps.
Visit our website to learn more.
************************************************************
