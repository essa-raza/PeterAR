from __future__ import annotations

import argparse
from pathlib import Path

from customer_notes_merger import default_output_path, merge_customer_notes


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Attach Commentaar/Mail/Whatsapp fields from an Excel file onto "
            "customer header rows in a semicolon-separated CSV, then save as XLSX."
        )
    )
    parser.add_argument("csv_file", type=Path, help="Input semicolon-separated CSV file.")
    parser.add_argument("excel_file", type=Path, help="Input Excel file containing columns H:J.")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output XLSX file path. Defaults to '<csv name> - merged.xlsx'.",
    )
    args = parser.parse_args()

    output_path = args.output or default_output_path(args.csv_file)
    result = merge_customer_notes(args.csv_file, args.excel_file, output_path)

    print(f"Created: {result.output_path}")
    print(f"Matched customer sections: {result.matched_count}")
    if result.unmatched_customers:
        print(f"Unmatched customer sections: {len(result.unmatched_customers)}")
        for name in result.unmatched_customers[:20]:
            print(f" - {name}")
        if len(result.unmatched_customers) > 20:
            print(f" ... and {len(result.unmatched_customers) - 20} more")


if __name__ == "__main__":
    main()
