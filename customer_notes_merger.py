from __future__ import annotations

import csv
import shutil
import tempfile
import time
import traceback
from collections import defaultdict, deque
from copy import copy
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Deque, Literal

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font


CSV_HEADERS = [
    "Type",
    "Omschrijving",
    "Datum",
    "Factuurnummer",
    "Vervaldatum",
    "Openstaand",
    "Bedrag",
    "Match status",
]
OUTPUT_HEADERS = CSV_HEADERS[:7] + ["Commentaar", "Mail", "Whatsapp"]
LEGACY_NOTE_COLUMNS_WITH_STATUS = (9, 10, 11)
LEGACY_NOTE_COLUMNS_NO_STATUS = (8, 9, 10)
CLOUD_PATH_MARKERS = (
    "onedrive",
    "zoho",
    "workdrive",
    "dropbox",
    "google drive",
    "icloud",
    "sharepoint",
)

ProcessingMode = Literal["auto", "sandbox", "direct"]


@dataclass
class LegacyRow:
    signature: tuple[str, ...]
    notes: tuple[object, object, object]
    row_fill: object | None


@dataclass
class MergeResult:
    output_path: Path
    matched_count: int
    unmatched_customers: list[str]
    total_customer_sections: int
    standard_log: list[str]
    debug_log: list[str]
    log_file_path: Path | None = None
    processing_mode_used: str = "direct"
    fallback_used: bool = False


class MergeLogger:
    def __init__(
        self,
        status_callback: Callable[[str], None] | None = None,
        debug_callback: Callable[[str], None] | None = None,
        debug_enabled: bool = False,
        save_log_file: bool = False,
    ) -> None:
        self.status_callback = status_callback
        self.debug_callback = debug_callback
        self.debug_enabled = debug_enabled
        self.save_log_file = save_log_file
        self.standard_lines: list[str] = []
        self.debug_lines: list[str] = []
        self._started_at = time.perf_counter()

    def _format_line(self, message: str) -> str:
        elapsed = time.perf_counter() - self._started_at
        return f"[{elapsed:7.2f}s] {message}"

    def status(self, message: str) -> None:
        if self.status_callback:
            self.status_callback(message)

    def standard(self, message: str, *, status_message: str | None = None) -> None:
        line = self._format_line(message)
        self.standard_lines.append(line)
        if status_message:
            self.status(status_message)

    def debug(self, message: str) -> None:
        line = self._format_line(message)
        self.debug_lines.append(line)
        if self.debug_enabled and self.debug_callback:
            self.debug_callback(line)

    def event(
        self,
        standard_message: str,
        *,
        status_message: str | None = None,
        debug_message: str | None = None,
    ) -> None:
        self.standard(standard_message, status_message=status_message)
        self.debug(debug_message or standard_message)

    def step_duration(self, label: str, started_at: float) -> None:
        self.debug(f"{label} completed in {time.perf_counter() - started_at:.2f}s")

    def debug_path(self, label: str, path: Path) -> None:
        resolved = path.expanduser().resolve(strict=False)
        marker = "cloud-synced path detected" if is_likely_cloud_path(resolved) else "local path"
        self.debug(f"{label}: {resolved} ({marker})")

    def debug_exception(self, context: str, exc: BaseException) -> None:
        self.debug(f"{context}: {exc}")
        self.debug(traceback.format_exc().rstrip())


def clean_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).replace("\r\n", "\n").replace("\r", "\n").strip()


def row_signature(values: list[object], width: int = 8) -> tuple[str, ...]:
    padded = list(values[:width])
    if len(padded) < width:
        padded.extend([""] * (width - len(padded)))
    return tuple(clean_text(value) for value in padded)


def is_likely_cloud_path(path: Path) -> bool:
    normalized = str(path).lower()
    return any(marker in normalized for marker in CLOUD_PATH_MARKERS)


def parse_csv_sections(csv_path: Path, logger: MergeLogger | None = None) -> tuple[list[str], list[list[str]]]:
    started_at = time.perf_counter()
    if logger:
        logger.event(
            "Customer CSV opened successfully.",
            status_message="Opening CSV file...",
            debug_message="Opening customer CSV for reading.",
        )
    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        sample = handle.read(4096)
        handle.seek(0)
        delimiter = ";" if sample.count(";") >= sample.count(",") else ","
        reader = csv.reader(handle, delimiter=delimiter)
        rows = list(reader)
    if logger:
        logger.debug(f"Detected '{delimiter}' as the CSV delimiter.")
        logger.step_duration("Reading customer CSV", started_at)

    if not rows:
        raise ValueError("CSV file is empty.")

    header = list(rows[0][:8])
    if header != CSV_HEADERS:
        raise ValueError(
            f"Unexpected CSV headers. Expected {CSV_HEADERS} but received {header}."
        )

    data_rows: list[list[str]] = []
    for row in rows[1:]:
        trimmed = list(row[:8])
        if len(trimmed) < 8:
            trimmed.extend([""] * (8 - len(trimmed)))
        data_rows.append(trimmed)

    if logger:
        logger.standard(f"Customer CSV rows loaded: {len(data_rows)}.")
    return header, data_rows


def build_legacy_row_map(
    xlsx_path: Path,
    logger: MergeLogger | None = None,
) -> tuple[dict[tuple[str, ...], Deque[LegacyRow]], int]:
    started_at = time.perf_counter()
    if logger:
        logger.event(
            "Source workbook opened successfully.",
            status_message="Opening source workbook...",
            debug_message="Opening source workbook for reading.",
        )
    workbook = load_workbook(xlsx_path)
    sheet = workbook.active
    header_values = [sheet.cell(1, col_idx).value for col_idx in range(1, sheet.max_column + 1)]
    normalized_headers = [clean_text(value).lower() for value in header_values]
    if normalized_headers[:11] == [value.lower() for value in CSV_HEADERS + ["Commentaar", "Mail", "Whatsapp"]]:
        note_columns = LEGACY_NOTE_COLUMNS_WITH_STATUS
        match_width = 8
    elif normalized_headers[:10] == [value.lower() for value in OUTPUT_HEADERS]:
        note_columns = LEGACY_NOTE_COLUMNS_NO_STATUS
        match_width = 7
    else:
        note_columns = LEGACY_NOTE_COLUMNS_WITH_STATUS
        match_width = 8

    if logger:
        logger.debug(
            f"Workbook matching mode uses {match_width} signature columns and note columns {note_columns}."
        )

    row_map: dict[tuple[str, ...], Deque[LegacyRow]] = defaultdict(deque)
    for row_idx in range(2, sheet.max_row + 1):
        signature = row_signature(
            [sheet.cell(row_idx, col_idx).value for col_idx in range(1, match_width + 1)],
            width=match_width,
        )
        notes = tuple(sheet.cell(row_idx, col_idx).value for col_idx in note_columns)

        row_fill = None
        for col_idx in range(1, min(sheet.max_column, 11) + 1):
            fill = sheet.cell(row_idx, col_idx).fill
            if fill and fill.fill_type:
                row_fill = copy(fill)
                break

        row_map[signature].append(LegacyRow(signature=signature, notes=notes, row_fill=row_fill))

    if logger:
        logger.event(
            f"Source workbook rows indexed: {sheet.max_row - 1}.",
            status_message="Reading workbook rows...",
            debug_message=f"Indexed {sheet.max_row - 1} legacy workbook rows for matching.",
        )
        logger.step_duration("Reading source workbook", started_at)
    return row_map, match_width


def default_output_path(csv_path: Path) -> Path:
    return csv_path.with_name(f"{csv_path.stem} - merged.xlsx")


def format_output_sheet(sheet) -> None:
    widths = {
        "A": 22,
        "B": 90,
        "C": 14,
        "D": 18,
        "E": 14,
        "F": 14,
        "G": 14,
        "H": 45,
        "I": 18,
        "J": 18,
    }
    wrap_alignment = Alignment(wrap_text=True, vertical="top")

    for column, width in widths.items():
        sheet.column_dimensions[column].width = width

    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = wrap_alignment

    for cell in sheet[1]:
        cell.font = Font(bold=True)

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:J{sheet.max_row}"


def write_output(
    output_path: Path,
    header: list[str],
    data_rows: list[list[str]],
    legacy_rows: dict[tuple[str, ...], Deque[LegacyRow]],
    match_width: int,
    logger: MergeLogger | None = None,
) -> tuple[int, list[str], int]:
    started_at = time.perf_counter()
    if logger:
        logger.event(
            "Output workbook build started.",
            status_message="Building merged workbook...",
            debug_message="Creating output workbook in memory.",
        )
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Merged"
    sheet.append(OUTPUT_HEADERS)

    matched = 0
    unmatched: list[str] = []

    for row_idx, row in enumerate(data_rows, start=2):
        padded = list(row)[: len(header)]
        if len(padded) < len(header):
            padded.extend([""] * (len(header) - len(padded)))

        legacy_row = None
        signature = row_signature(padded, width=match_width)
        matches = legacy_rows.get(signature)
        if matches:
            legacy_row = matches.popleft()
            matched += 1
        else:
            unmatched.append(" | ".join(signature))

        notes = list(legacy_row.notes) if legacy_row else ["", "", ""]
        output_row = padded[:7] + notes
        sheet.append(output_row)

        if legacy_row and legacy_row.row_fill:
            for col_idx in range(1, 11):
                sheet.cell(row_idx, col_idx).fill = copy(legacy_row.row_fill)

    if logger:
        logger.event(
            f"Merge rows prepared. Matched: {matched}. Unmatched: {len(unmatched)}.",
            status_message="Merging notes into output rows...",
            debug_message=(
                f"Prepared {len(data_rows)} output rows with {matched} matched rows "
                f"and {len(unmatched)} unmatched rows."
            ),
        )
        logger.event(
            "Workbook formatting applied.",
            status_message="Formatting workbook...",
            debug_message="Applying workbook formatting.",
        )
    format_output_sheet(sheet)
    if logger:
        logger.event(
            "Output workbook save started.",
            status_message="Saving output workbook...",
            debug_message="Saving merged workbook to disk.",
        )
    workbook.save(output_path)
    if logger:
        logger.standard(f"Output workbook saved: {output_path.name}.")
        logger.step_duration("Writing output workbook", started_at)
    return matched, unmatched, len(data_rows)


def write_log_file(log_file_path: Path, standard_lines: list[str], debug_lines: list[str]) -> None:
    sections = ["Standard log", "-" * 12, *standard_lines]
    if debug_lines:
        sections.extend(["", "Debug details", "-" * 13, *debug_lines])
    log_file_path.write_text("\n".join(sections) + "\n", encoding="utf-8")


def run_merge_pipeline(
    csv_file: Path,
    excel_file: Path,
    output_file: Path,
    logger: MergeLogger,
) -> tuple[int, list[str], int]:
    header, data_rows = parse_csv_sections(csv_file, logger=logger)
    legacy_rows, match_width = build_legacy_row_map(excel_file, logger=logger)
    return write_output(output_file, header, data_rows, legacy_rows, match_width, logger=logger)


def process_with_sandbox(
    csv_file: Path,
    excel_file: Path,
    output_file: Path,
    logger: MergeLogger,
) -> tuple[int, list[str], int]:
    stage_started_at = time.perf_counter()
    with tempfile.TemporaryDirectory(prefix="customer-notes-merger-") as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        local_csv = temp_dir / csv_file.name
        local_excel = temp_dir / excel_file.name
        local_output = temp_dir / output_file.name

        logger.event(
            "Local sandbox workspace created.",
            status_message="Preparing local sandbox...",
            debug_message=f"Using sandbox directory: {temp_dir}",
        )
        logger.debug_path("Original customer CSV path", csv_file)
        logger.debug_path("Original source workbook path", excel_file)
        logger.debug_path("Requested output path", output_file)
        logger.debug_path("Sandbox customer CSV path", local_csv)
        logger.debug_path("Sandbox source workbook path", local_excel)
        logger.debug_path("Sandbox output path", local_output)

        copy_started_at = time.perf_counter()
        shutil.copy2(csv_file, local_csv)
        shutil.copy2(excel_file, local_excel)
        logger.event(
            "Input files copied to local sandbox.",
            status_message="Copying input files locally...",
            debug_message="Copied input files from source paths into the sandbox workspace.",
        )
        logger.step_duration("Copying input files into sandbox", copy_started_at)

        matched, unmatched, total = run_merge_pipeline(local_csv, local_excel, local_output, logger)

        move_started_at = time.perf_counter()
        output_file.parent.mkdir(parents=True, exist_ok=True)
        if output_file.exists():
            output_file.unlink()
        shutil.move(str(local_output), str(output_file))
        logger.event(
            "Sandbox output moved to the requested destination.",
            status_message="Finalizing output file...",
            debug_message="Moved the finished workbook from the sandbox back to the requested destination.",
        )
        logger.step_duration("Moving output file back to destination", move_started_at)
        logger.step_duration("Sandbox processing", stage_started_at)
        return matched, unmatched, total


def process_directly(
    csv_file: Path,
    excel_file: Path,
    output_file: Path,
    logger: MergeLogger,
) -> tuple[int, list[str], int]:
    logger.event(
        "Direct file access mode started.",
        status_message="Processing files directly...",
        debug_message="Running merge directly against the selected file paths.",
    )
    logger.debug_path("Direct customer CSV path", csv_file)
    logger.debug_path("Direct source workbook path", excel_file)
    logger.debug_path("Direct output path", output_file)
    return run_merge_pipeline(csv_file, excel_file, output_file, logger)


def merge_customer_notes(
    csv_path: Path | str,
    excel_path: Path | str,
    output_path: Path | str | None = None,
    status_callback: Callable[[str], None] | None = None,
    debug_callback: Callable[[str], None] | None = None,
    debug_mode: bool = False,
    save_log_file: bool = False,
    processing_mode: ProcessingMode = "auto",
) -> MergeResult:
    overall_started_at = time.perf_counter()
    csv_file = Path(csv_path)
    excel_file = Path(excel_path)
    output_file = Path(output_path) if output_path else default_output_path(csv_file)
    logger = MergeLogger(
        status_callback=status_callback,
        debug_callback=debug_callback,
        debug_enabled=debug_mode,
        save_log_file=save_log_file,
    )

    logger.event(
        "Merge started.",
        status_message="Starting merge...",
        debug_message="Starting merge process.",
    )

    matched = 0
    unmatched: list[str] = []
    total = 0
    processing_mode_used = "direct"
    fallback_used = False

    should_try_sandbox = processing_mode in {"auto", "sandbox"}
    if should_try_sandbox:
        try:
            matched, unmatched, total = process_with_sandbox(csv_file, excel_file, output_file, logger)
            processing_mode_used = "sandbox"
        except Exception as exc:  # noqa: BLE001
            logger.standard(
                "Sandbox processing failed.",
                status_message="Sandbox processing failed. Switching to direct mode...",
            )
            logger.debug_exception("Sandbox processing failed", exc)
            if processing_mode == "sandbox":
                raise
            fallback_used = True
            matched, unmatched, total = process_directly(csv_file, excel_file, output_file, logger)
            processing_mode_used = "direct"
    else:
        matched, unmatched, total = process_directly(csv_file, excel_file, output_file, logger)
        processing_mode_used = "direct"

    logger.event(
        f"Merge completed successfully in {time.perf_counter() - overall_started_at:.2f}s.",
        status_message="Merge completed successfully.",
        debug_message=(
            f"Merge completed successfully in {time.perf_counter() - overall_started_at:.2f}s "
            f"using {processing_mode_used} mode."
        ),
    )

    log_file_path: Path | None = None
    if save_log_file:
        log_file_path = output_file.with_suffix(".log")
        logger.standard(f"Log file written: {log_file_path.name}.")
        if debug_mode:
            logger.debug_path("Log file path", log_file_path)
        write_log_file(log_file_path, logger.standard_lines, logger.debug_lines if debug_mode else [])

    return MergeResult(
        output_path=output_file,
        matched_count=matched,
        unmatched_customers=unmatched,
        total_customer_sections=total,
        standard_log=logger.standard_lines,
        debug_log=logger.debug_lines,
        log_file_path=log_file_path,
        processing_mode_used=processing_mode_used,
        fallback_used=fallback_used,
    )
