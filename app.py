from __future__ import annotations

import argparse
import threading
from pathlib import Path
from typing import Sequence
from tkinter import filedialog, messagebox

import customtkinter as ctk

from customer_notes_merger import default_output_path, merge_customer_notes


APP_TITLE = "Customer Notes Merger"


class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("green")
        ctk.set_widget_scaling(0.92)
        ctk.set_window_scaling(0.92)

        self.title(APP_TITLE)
        self.geometry("1080x720")
        self.minsize(980, 660)

        self.csv_var = ctk.StringVar()
        self.excel_var = ctk.StringVar()
        self.output_var = ctk.StringVar()
        self.status_var = ctk.StringVar(value="Choose the CSV and Excel files to begin.")
        self.summary_var = ctk.StringVar(value="No merge has been run yet.")

        self._build_ui()

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        shell = ctk.CTkFrame(self, fg_color="#f4f1e8", corner_radius=0)
        shell.grid(row=0, column=0, sticky="nsew")
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(1, weight=1)

        hero = ctk.CTkFrame(shell, fg_color="#18392b", corner_radius=28)
        hero.grid(row=0, column=0, padx=24, pady=(24, 18), sticky="ew")
        hero.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            hero,
            text="Merge follow-up notes into the customer CSV",
            font=ctk.CTkFont(family="Segoe UI", size=30, weight="bold"),
            text_color="#f7f4ea",
        ).grid(row=0, column=0, padx=28, pady=(26, 8), sticky="w")

        ctk.CTkLabel(
            hero,
            text=(
                "Transfer Commentaar, Mail, and Whatsapp details from the previous workbook, "
                "keep Match status in column H, and export the result as a polished Excel workbook."
            ),
            justify="left",
            wraplength=760,
            font=ctk.CTkFont(family="Segoe UI", size=15),
            text_color="#d7ead8",
        ).grid(row=1, column=0, padx=28, pady=(0, 24), sticky="w")

        body = ctk.CTkFrame(shell, fg_color="transparent")
        body.grid(row=1, column=0, padx=24, pady=(0, 24), sticky="nsew")
        body.grid_columnconfigure(0, weight=3)
        body.grid_columnconfigure(1, weight=2)
        body.grid_rowconfigure(0, weight=1)

        form_card = ctk.CTkFrame(body, fg_color="#fbf8f0", corner_radius=24, border_width=1, border_color="#d9d0bc")
        form_card.grid(row=0, column=0, padx=(0, 12), sticky="nsew")
        form_card.grid_columnconfigure(0, weight=1)

        sidebar = ctk.CTkFrame(body, fg_color="#dce8dd", corner_radius=24)
        sidebar.grid(row=0, column=1, padx=(12, 0), sticky="nsew")
        sidebar.grid_columnconfigure(0, weight=1)

        self._add_path_picker(
            form_card,
            row=0,
            title="Customer CSV",
            description="Select the semicolon-separated CSV export from Yuki.",
            variable=self.csv_var,
            command=self._browse_csv,
            button_text="Browse CSV",
        )
        self._add_path_picker(
            form_card,
            row=1,
            title="Source Excel",
            description="Select the workbook that contains columns H to J.",
            variable=self.excel_var,
            command=self._browse_excel,
            button_text="Browse Excel",
        )
        self._add_path_picker(
            form_card,
            row=2,
            title="Output Workbook",
            description="Choose where the merged Excel file should be saved.",
            variable=self.output_var,
            command=self._browse_output,
            button_text="Save As",
        )

        action_bar = ctk.CTkFrame(form_card, fg_color="transparent")
        action_bar.grid(row=3, column=0, padx=26, pady=(8, 12), sticky="ew")
        action_bar.grid_columnconfigure(0, weight=1)

        self.merge_button = ctk.CTkButton(
            action_bar,
            text="Create Merged Workbook",
            height=52,
            corner_radius=16,
            fg_color="#1e5c41",
            hover_color="#163f2d",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            command=self._start_merge,
        )
        self.merge_button.grid(row=0, column=0, sticky="ew")

        status_card = ctk.CTkFrame(form_card, fg_color="#efe7d8", corner_radius=18)
        status_card.grid(row=4, column=0, padx=26, pady=(4, 26), sticky="ew")
        status_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            status_card,
            text="Status",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color="#4b4336",
        ).grid(row=0, column=0, padx=18, pady=(16, 4), sticky="w")

        ctk.CTkLabel(
            status_card,
            textvariable=self.status_var,
            justify="left",
            wraplength=560,
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color="#5c5347",
        ).grid(row=1, column=0, padx=18, pady=(0, 16), sticky="w")

        ctk.CTkLabel(
            sidebar,
            text="What this app does",
            font=ctk.CTkFont(family="Segoe UI", size=24, weight="bold"),
            text_color="#173226",
        ).grid(row=0, column=0, padx=24, pady=(26, 12), sticky="w")

        bullets = [
            "Matches rows between the new CSV and the previous workbook using columns A to H.",
            "Collects Commentaar, Mail, and Whatsapp values from columns I to K.",
            "Preserves row highlighting by coloring the full output row when a matched row was colored.",
            "Exports a ready-to-share Excel file without changing the source files.",
        ]
        for index, bullet in enumerate(bullets, start=1):
            ctk.CTkLabel(
                sidebar,
                text=f"{index}. {bullet}",
                justify="left",
                wraplength=280,
                font=ctk.CTkFont(family="Segoe UI", size=15),
                text_color="#274637",
            ).grid(row=index, column=0, padx=24, pady=6, sticky="w")

        summary_card = ctk.CTkFrame(sidebar, fg_color="#f8f4ea", corner_radius=20)
        summary_card.grid(row=5, column=0, padx=20, pady=(18, 20), sticky="ew")
        summary_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            summary_card,
            text="Latest run",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color="#483f34",
        ).grid(row=0, column=0, padx=18, pady=(16, 8), sticky="w")

        ctk.CTkLabel(
            summary_card,
            textvariable=self.summary_var,
            justify="left",
            wraplength=270,
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color="#5f584f",
        ).grid(row=1, column=0, padx=18, pady=(0, 18), sticky="w")

    def _add_path_picker(
        self,
        parent: ctk.CTkFrame,
        row: int,
        title: str,
        description: str,
        variable: ctk.StringVar,
        command,
        button_text: str,
    ) -> None:
        block = ctk.CTkFrame(parent, fg_color="transparent")
        block.grid(row=row, column=0, padx=26, pady=(26 if row == 0 else 10, 4), sticky="ew")
        block.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            block,
            text=title,
            font=ctk.CTkFont(family="Segoe UI", size=19, weight="bold"),
            text_color="#1c362a",
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            block,
            text=description,
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color="#6b6256",
        ).grid(row=1, column=0, pady=(4, 10), sticky="w")

        row_frame = ctk.CTkFrame(block, fg_color="transparent")
        row_frame.grid(row=2, column=0, sticky="ew")
        row_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkEntry(
            row_frame,
            textvariable=variable,
            height=46,
            corner_radius=14,
            fg_color="#fffdf7",
            border_color="#d3cab6",
            text_color="#2f2a25",
            font=ctk.CTkFont(family="Segoe UI", size=14),
        ).grid(row=0, column=0, padx=(0, 10), sticky="ew")

        ctk.CTkButton(
            row_frame,
            text=button_text,
            width=132,
            height=46,
            corner_radius=14,
            fg_color="#e2dccd",
            hover_color="#d5cebd",
            text_color="#2d2a26",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            command=command,
        ).grid(row=0, column=1, sticky="e")

    def _browse_csv(self) -> None:
        path = filedialog.askopenfilename(
            title="Select CSV file",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            self.csv_var.set(path)
            if not self.output_var.get():
                self.output_var.set(str(default_output_path(Path(path))))

    def _browse_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.excel_var.set(path)

    def _browse_output(self) -> None:
        initial = self.output_var.get() or "merged-output.xlsx"
        path = filedialog.asksaveasfilename(
            title="Choose output file",
            defaultextension=".xlsx",
            initialfile=Path(initial).name,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.output_var.set(path)

    def _start_merge(self) -> None:
        csv_path = self.csv_var.get().strip()
        excel_path = self.excel_var.get().strip()
        output_path = self.output_var.get().strip()

        if not csv_path or not excel_path:
            messagebox.showerror(APP_TITLE, "Please choose both the CSV file and the Excel file.")
            return

        if not output_path:
            output_path = str(default_output_path(Path(csv_path)))
            self.output_var.set(output_path)

        self.merge_button.configure(state="disabled")
        self.status_var.set("Processing files and building the merged workbook...")
        self.summary_var.set("Running merge...")

        worker = threading.Thread(
            target=self._run_merge,
            args=(csv_path, excel_path, output_path),
            daemon=True,
        )
        worker.start()

    def _run_merge(self, csv_path: str, excel_path: str, output_path: str) -> None:
        try:
            result = merge_customer_notes(csv_path, excel_path, output_path)
        except Exception as exc:  # noqa: BLE001
            self.after(0, lambda: self._show_error(str(exc)))
            return

        self.after(0, lambda: self._show_success(result))

    def _show_success(self, result) -> None:
        self.merge_button.configure(state="normal")
        unmatched_count = len(result.unmatched_customers)
        self.status_var.set(
            f"Merged workbook created successfully at:\n{result.output_path}"
        )
        self.summary_var.set(
            f"Matched {result.matched_count} of {result.total_customer_sections} rows.\n"
            f"Unmatched rows: {unmatched_count}"
        )
        messagebox.showinfo(
            APP_TITLE,
            (
                f"Done.\n\nOutput: {result.output_path}\n"
                f"Matched rows: {result.matched_count}\n"
                f"Unmatched rows: {unmatched_count}"
            ),
        )

    def _show_error(self, message: str) -> None:
        self.merge_button.configure(state="normal")
        self.status_var.set("The merge could not be completed. Please review the file paths and try again.")
        self.summary_var.set(f"Last error:\n{message}")
        messagebox.showerror(APP_TITLE, message)


def run_cli(csv_path: str, excel_path: str, output_path: str | None) -> int:
    result = merge_customer_notes(
        csv_path,
        excel_path,
        output_path or str(default_output_path(Path(csv_path))),
    )
    print(f"Created: {result.output_path}")
    print(f"Matched rows: {result.matched_count}")
    print(f"Unmatched rows: {len(result.unmatched_customers)}")
    return 0


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--cli", action="store_true")
    parser.add_argument("--csv")
    parser.add_argument("--excel")
    parser.add_argument("--output")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    if args.cli:
        if not args.csv or not args.excel:
            raise SystemExit("CLI mode requires --csv and --excel.")
        return run_cli(args.csv, args.excel, args.output)

    app = App()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
