# Customer Notes Merger

Please choose the version that matches your computer:

- `CustomerNotesMerger-Windows.exe`
  Use this for Windows PCs.

- `CustomerNotesMerger-Mac-Intel.dmg`
  Use this for older Macs with an Intel processor.

- `CustomerNotesMerger-Mac-Apple-Silicon.dmg`
  Use this for newer Macs with Apple Silicon chips such as `M1`, `M2`, `M3`, or `M4`.

## How to check which Mac you have

1. Click the Apple menu.
2. Click `About This Mac`.
3. Check the processor or chip:

- If it says `Intel`, use the Intel version.
- If it says `Apple M1`, `M2`, `M3`, or `M4`, use the Apple Silicon version.

## What the app does

The app takes:

- one CSV file
- one Excel file

It copies the information from Excel columns `H` to `J` to the matching customer row in the CSV data and then saves the result as a new Excel file.

It also:

- combines repeated comment lines into one field
- saves the final result as `.xlsx`
- formats the comments column so the text wraps inside the cells

## Notes

- The original files are not changed.
- A new output Excel file is created.
- If a customer exists in the CSV but not in the Excel file, that customer will remain without added values.
