from __future__ import annotations

import csv
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment


LEGAL_TOKENS = {
    "BV",
    "BVBA",
    "SRL",
    "SPRL",
    "NV",
    "CV",
    "CVBA",
    "COMMV",
    "VOF",
}


@dataclass
class CustomerInfo:
    comments: list[str] = field(default_factory=list)
    mails: list[str] = field(default_factory=list)
    whatsapps: list[str] = field(default_factory=list)

    def add(self, comment: object, mail: object, whatsapp: object) -> None:
        if comment not in (None, ""):
            self.comments.append(clean_text(comment))
        if mail not in (None, ""):
            self.mails.append(clean_text(mail))
        if whatsapp not in (None, ""):
            self.whatsapps.append(clean_text(whatsapp))

    def merged(self) -> tuple[str, str, str]:
        return (
            join_unique(self.comments),
            join_unique(self.mails),
            join_unique(self.whatsapps),
        )


@dataclass
class MergeResult:
    output_path: Path
    matched_count: int
    unmatched_customers: list[str]
    total_customer_sections: int


def clean_text(value: object) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")

    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.strip() for line in text.split("\n")]
    return "\n".join(line for line in lines if line)


def join_unique(values: Iterable[str]) -> str:
    seen: set[str] = set()
    ordered: list[str] = []
    for value in values:
        cleaned = clean_text(value)
        if cleaned and cleaned not in seen:
            seen.add(cleaned)
            ordered.append(cleaned)
    return "\n".join(ordered)


def normalize_name(name: str) -> str:
    text = clean_text(name).upper()
    text = re.sub(r"\([^)]*\)", " ", text)
    text = text.replace("&", " AND ")
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    tokens = [token for token in text.split() if token and token not in LEGAL_TOKENS]
    return " ".join(tokens)


def alias_candidates(name: str) -> list[str]:
    base = normalize_name(name)
    aliases = [base]

    stripped_prefix = re.sub(
        r"^(?:BV|BVBA|SRL|SPRL|NV|CV|CVBA|COMMV|VOF)\s+",
        "",
        clean_text(name),
        flags=re.IGNORECASE,
    )
    stripped_normalized = normalize_name(stripped_prefix)
    if stripped_normalized and stripped_normalized not in aliases:
        aliases.append(stripped_normalized)

    compact = base.replace(" ", "")
    if compact and compact not in aliases:
        aliases.append(compact)

    return aliases


def parse_csv_sections(csv_path: Path) -> tuple[list[str], list[list[str]]]:
    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle, delimiter=";")
        rows = list(reader)

    if not rows:
        raise ValueError("CSV file is empty.")

    header = list(rows[0])
    data_rows = [list(row) for row in rows[1:]]
    return header, data_rows


def is_excel_customer_header(row_values: list[object]) -> bool:
    first = row_values[0]
    if first in (None, ""):
        return False
    return all(value in (None, "") for value in row_values[1:7])


def build_customer_map(xlsx_path: Path) -> dict[str, CustomerInfo]:
    workbook = load_workbook(xlsx_path, data_only=True)
    sheet = workbook.active

    customer_map: dict[str, CustomerInfo] = {}
    current_names: list[str] = []

    for row_idx in range(2, sheet.max_row + 1):
        row = [sheet.cell(row_idx, col_idx).value for col_idx in range(1, 11)]

        if is_excel_customer_header(row):
            header_name = clean_text(row[0])
            current_names = alias_candidates(header_name)
            info = customer_map.setdefault(current_names[0], CustomerInfo())
            info.add(row[7], row[8], row[9])
            for alias in current_names[1:]:
                customer_map[alias] = info
            continue

        if not current_names:
            continue

        comment, mail, whatsapp = row[7], row[8], row[9]
        if comment in (None, "") and mail in (None, "") and whatsapp in (None, ""):
            continue

        customer_map[current_names[0]].add(comment, mail, whatsapp)

    return customer_map


def default_output_path(csv_path: Path) -> Path:
    return csv_path.with_name(f"{csv_path.stem} - merged.xlsx")


def format_output_sheet(sheet) -> None:
    # Column I contains the merged comments field in the exported workbook.
    sheet.column_dimensions["I"].width = 42
    for cell in sheet["I"]:
        cell.alignment = Alignment(wrap_text=True, vertical="top")


def write_output(
    output_path: Path,
    header: list[str],
    data_rows: list[list[str]],
    customer_map: dict[str, CustomerInfo],
) -> tuple[int, list[str], int]:
    final_header = list(header) + ["Commentaar", "Mail", "Whatsapp"]
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Merged"
    sheet.append(final_header)

    matched = 0
    unmatched: list[str] = []
    total_customer_sections = 0

    for row in data_rows:
        padded = list(row) + [""] * max(0, len(header) - len(row))
        customer_name = padded[0].strip() if padded else ""
        match = None

        is_customer_header = (
            customer_name and all(not value.strip() for value in padded[1:len(header)])
        )
        if is_customer_header:
            total_customer_sections += 1
            normalized = normalize_name(customer_name)
            match = customer_map.get(normalized)
            if match is None:
                match = customer_map.get(normalized.replace(" ", ""))

            if match is not None:
                matched += 1
            else:
                unmatched.append(customer_name)

        extras = list(match.merged()) if match else ["", "", ""]
        sheet.append(padded + extras)

    format_output_sheet(sheet)
    workbook.save(output_path)
    return matched, unmatched, total_customer_sections


def merge_customer_notes(
    csv_path: Path | str,
    excel_path: Path | str,
    output_path: Path | str | None = None,
) -> MergeResult:
    csv_file = Path(csv_path)
    excel_file = Path(excel_path)
    output_file = Path(output_path) if output_path else default_output_path(csv_file)

    header, data_rows = parse_csv_sections(csv_file)
    customer_map = build_customer_map(excel_file)
    matched, unmatched, total = write_output(output_file, header, data_rows, customer_map)

    return MergeResult(
        output_path=output_file,
        matched_count=matched,
        unmatched_customers=unmatched,
        total_customer_sections=total,
    )
