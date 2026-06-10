#!/usr/bin/env python3
"""Desktop GUI for the WSJ Treasury / Barchart Futures scraper."""

from __future__ import annotations

import sys
import threading
import traceback
from pathlib import Path
from tkinter import (
    BOTH, END, LEFT, RIGHT, X, Y,
    Button, Entry, Frame, Label, StringVar, Text, Tk,
    filedialog, messagebox,
)
from tkinter import ttk

# Resolve the correct base directory whether running frozen or as a .py file
if getattr(sys, "frozen", False):
    # PyInstaller bundle: _MEIPASS is the temp extraction dir (onefile)
    # or the app dir (onedir). We want the folder beside the actual executable.
    HERE = Path(sys.executable).resolve().parent
    _bundle = Path(getattr(sys, "_MEIPASS", HERE))
    # Add both so the scraper module can be found either way
    for _p in (str(HERE), str(_bundle)):
        if _p not in sys.path:
            sys.path.insert(0, _p)
else:
    HERE = Path(__file__).resolve().parent
    if str(HERE) not in sys.path:
        sys.path.insert(0, str(HERE))

try:
    import wsj_treasury_scraper_fixed as scraper
except Exception as exc:
    scraper = None
    _import_error = exc
else:
    _import_error = None

WSJ_URL      = "https://www.wsj.com/market-data/bonds/treasuries#treasuryB"
BARCHART_URL = "https://www.barchart.com/futures/financials?viewName=main"

SOURCE_WSJ      = "WSJ Treasuries"
SOURCE_BARCHART = "Barchart Futures"
SOURCES = [SOURCE_WSJ, SOURCE_BARCHART]


class TreasuryApp(Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Treasury / Futures Scraper")
        self.geometry("1220x800")
        self.minsize(1000, 660)
        # Prevent the window from closing mid-scrape accidentally
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.source_var    = StringVar(value=SOURCE_WSJ)
        self.url_var       = StringVar(value=WSJ_URL)
        self.output_var    = StringVar(value=str(HERE / "wsj_treasuries.xlsx"))
        self.reference_var = StringVar(value="")
        self.status_var    = StringVar(value="Ready.")

        self._worker: threading.Thread | None = None
        self._current_df = None

        self._build_ui()
        if _import_error is not None:
            self._log(f"Import warning: {_import_error}")
            self._set_status("The scraper module could not be imported.")

    # ── UI ─────────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        outer = Frame(self, padx=12, pady=12)
        outer.pack(fill=BOTH, expand=True)

        top = Frame(outer)
        top.pack(fill=X)

        # Row 0: source selector
        Label(top, text="Data source").grid(row=0, column=0, sticky="w", pady=(0, 4))
        src_frame = Frame(top)
        src_frame.grid(row=0, column=1, sticky="w", padx=(8, 8), pady=(0, 4))
        for src in SOURCES:
            ttk.Radiobutton(
                src_frame, text=src, variable=self.source_var,
                value=src, command=self._on_source_changed,
            ).pack(side=LEFT, padx=(0, 16))

        # Row 1: URL
        Label(top, text="URL").grid(row=1, column=0, sticky="w", pady=(0, 4))
        Entry(top, textvariable=self.url_var).grid(
            row=1, column=1, sticky="we", padx=(8, 8), pady=(0, 4))

        # Row 2: output file
        Label(top, text="Output file").grid(row=2, column=0, sticky="w", pady=(0, 4))
        out_row = Frame(top)
        out_row.grid(row=2, column=1, sticky="we", padx=(8, 8), pady=(0, 4))
        Entry(out_row, textvariable=self.output_var).pack(side=LEFT, fill=X, expand=True)
        Button(out_row, text="Browse…", command=self._browse_output).pack(side=RIGHT, padx=(8, 0))

        # Row 3: reference date (WSJ only)
        self._ref_label = Label(top, text="Reference date (YYYY-MM-DD)")
        self._ref_label.grid(row=3, column=0, sticky="w", pady=(0, 4))
        self._ref_entry = Entry(top, textvariable=self.reference_var)
        self._ref_entry.grid(row=3, column=1, sticky="we", padx=(8, 8), pady=(0, 4))

        # Row 4: action buttons
        btn_row = Frame(top)
        btn_row.grid(row=4, column=1, sticky="w", padx=(8, 8), pady=(6, 2))
        Button(btn_row, text="Run scrape",         command=self._run_clicked,        width=16).pack(side=LEFT)
        Button(btn_row, text="Open output folder", command=self._open_output_folder, width=18).pack(side=LEFT, padx=(8, 0))

        # Row 5: log buttons on separate row so they're never hidden
        log_btn_row = Frame(top)
        log_btn_row.grid(row=5, column=1, sticky="w", padx=(8, 8), pady=(2, 8))
        Button(log_btn_row, text="Save log…", command=self._save_log,  width=14).pack(side=LEFT)
        Button(log_btn_row, text="Clear log", command=self._clear_log, width=12).pack(side=LEFT, padx=(8, 0))

        top.columnconfigure(1, weight=1)

        # ── Middle: preview + log ──────────────────────────────────────────────
        mid = Frame(outer)
        mid.pack(fill=BOTH, expand=True, pady=(8, 8))

        preview_box = Frame(mid)
        preview_box.pack(side=LEFT, fill=BOTH, expand=True)
        Label(preview_box, text="Preview").pack(anchor="w")

        self.tree = ttk.Treeview(preview_box, show="headings")
        yscroll = ttk.Scrollbar(preview_box, orient="vertical",   command=self.tree.yview)
        xscroll = ttk.Scrollbar(preview_box, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.pack(fill=BOTH, expand=True, side=LEFT)
        yscroll.pack(side=RIGHT, fill=Y)
        xscroll.pack(side="bottom", fill=X)

        log_box = Frame(mid, width=420)
        log_box.pack(side=RIGHT, fill=Y, padx=(12, 0))
        log_box.pack_propagate(False)   # keep fixed width
        Label(log_box, text="Log").pack(anchor="w")
        log_scroll = ttk.Scrollbar(log_box, orient="vertical")
        self.log = Text(log_box, wrap="word", yscrollcommand=log_scroll.set)
        log_scroll.config(command=self.log.yview)
        log_scroll.pack(side=RIGHT, fill=Y)
        self.log.pack(fill=BOTH, expand=True)

        # ── Status bar ─────────────────────────────────────────────────────────
        bottom = Frame(outer)
        bottom.pack(fill=X)
        ttk.Separator(bottom, orient="horizontal").pack(fill=X, pady=(0, 6))
        Label(bottom, textvariable=self.status_var, anchor="w").pack(fill=X)

        self._log("Ready. Choose a data source, set an output file, and press Run scrape.")
        self._on_source_changed()

    # ── Helpers ────────────────────────────────────────────────────────────────

    def _on_close(self) -> None:
        if self._worker and self._worker.is_alive():
            if not messagebox.askyesno("Scrape running",
                    "A scrape is still running. Quit anyway?"):
                return
        self.destroy()

    def _on_source_changed(self) -> None:
        src = self.source_var.get()
        if src == SOURCE_WSJ:
            self.url_var.set(WSJ_URL)
            self.output_var.set(str(HERE / "wsj_treasuries.xlsx"))
            self._ref_label.grid()
            self._ref_entry.grid()
        else:
            self.url_var.set(BARCHART_URL)
            self.output_var.set(str(HERE / "barchart_futures.xlsx"))
            self._ref_label.grid_remove()
            self._ref_entry.grid_remove()

    def _set_status(self, text: str) -> None:
        self.status_var.set(text)

    def _log(self, text: str) -> None:
        self.log.insert(END, text.rstrip() + "\n")
        self.log.see(END)

    def _clear_log(self) -> None:
        self.log.delete("1.0", END)
        self._log("Log cleared.")

    def _save_log(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save log",
            defaultextension=".txt",
            filetypes=[("Text file", "*.txt"), ("All files", "*.*")],
            initialfile="scraper_log.txt",
            initialdir=str(HERE),
        )
        if path:
            try:
                Path(path).write_text(self.log.get("1.0", END), encoding="utf-8")
                self._log(f"Log saved to {path}")
            except Exception as exc:
                messagebox.showerror("Save failed", str(exc))

    def _browse_output(self) -> None:
        filename = filedialog.asksaveasfilename(
            title="Choose output Excel file",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")],
            initialfile=Path(self.output_var.get()).name or "output.xlsx",
            initialdir=str(HERE),
        )
        if filename:
            self.output_var.set(filename)

    def _open_output_folder(self) -> None:
        path   = Path(self.output_var.get()).expanduser()
        folder = path.parent if path.suffix else path
        try:
            if sys.platform.startswith("win"):
                import os; os.startfile(str(folder))
            elif sys.platform == "darwin":
                import subprocess; subprocess.run(["open",     str(folder)], check=False)
            else:
                import subprocess; subprocess.run(["xdg-open", str(folder)], check=False)
        except Exception as exc:
            messagebox.showerror("Open folder failed", str(exc))

    def _run_clicked(self) -> None:
        if self._worker and self._worker.is_alive():
            messagebox.showinfo("Busy", "A scrape is already running.")
            return
        if scraper is None:
            messagebox.showerror("Import error",
                f"Could not import wsj_treasury_scraper: {_import_error}")
            return

        src = self.source_var.get()
        self._set_status(f"Running scrape ({src})…")
        self._log(f"─── Starting {src} scrape ───")
        target = self._run_barchart if src == SOURCE_BARCHART else self._run_wsj
        self._worker = threading.Thread(target=target, daemon=True)
        self._worker.start()

    # ── WSJ scrape ─────────────────────────────────────────────────────────────

    def _run_wsj(self) -> None:
        try:
            url    = self.url_var.get().strip()
            output = Path(self.output_var.get().strip()).expanduser()

            ref_raw = self.reference_var.get().strip()
            reference_date = scraper.get_reference_date(url, ref_raw or None)

            self.after(0, lambda: self._log(f"Reference date: {reference_date.isoformat()}"))
            self.after(0, lambda: self._log("Scraping Treasury table…"))

            df = scraper.scrape_treasuries(url)
            df = scraper.add_payment_columns(df, reference_date)

            output.parent.mkdir(parents=True, exist_ok=True)
            formula_rows = 14
            sheet_name   = "Treasuries"

            with scraper.pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=formula_rows)
                ws = writer.book[sheet_name]
                ws["A1"]  = "Treasury Formulas"
                ws["A2"]  = "Coupon Payment";    ws["B2"]  = "Coupon/2*1000"
                ws["A3"]  = "PV0";               ws["B3"]  = "PV(Ask Yield/2, Number of Payments Until Maturity, Coupon Payment, 1000)"
                ws["A4"]  = "PV1";               ws["B4"]  = "PV((Ask Yield+1%)/2, Number of Payments Until Maturity-2, Coupon Payment, 1000)"
                ws["A5"]  = "P0";                ws["B5"]  = "PV0 * (1+Asked Yield/2)^(Days Since Last Payment/182)"
                ws["A6"]  = "P1";                ws["B6"]  = "PV1 * (1+(Asked Yield+1%)/2)^(Days Since Last Payment/182)"
                ws["A7"]  = "Simulated Return";  ws["B7"]  = "(P1- P0 + Coupon*10)/P0"
                ws["A8"]  = "MACAULAY DURATION"; ws["B8"]  = "DURATION(Current Date, Maturity Date, Coupon Rate, Ask Yield, 2)"
                ws["A9"]  = "MODIFIED DURATION"; ws["B9"]  = "MDURATION(Current Date, Maturity Date, Coupon Rate, Ask Yield, 2)"
                ws["A10"] = "PVUP";              ws["B10"] = "PV(Ask Yield+1%/2, Number of Payments Until Maturity, Coupon Payment, 1000)"
                ws["A11"] = "PVDN";              ws["B11"] = "PV(Ask Yield-1%/2, Number of Payments Until Maturity, Coupon Payment, 1000)"
                ws["A12"] = "PUP";               ws["B12"] = "PVUP * (1+(Asked Yield+1%)/2)^(Days Since Last Payment/182)"
                ws["A13"] = "PDN";               ws["B13"] = "PVDN * (1+(Asked Yield-1%)/2)^(Days Since Last Payment/182)"
                ws["A14"] = "EFDURATION";        ws["B14"] = "(PDN- PUP)/(2*.01*P0)"

            n = len(df)
            self.after(0, lambda: self._update_preview(df))
            self.after(0, lambda: self._log(f"✓ Saved {n} rows → {output.resolve()}"))
            self.after(0, lambda: self._set_status(f"Done — {n} rows saved to {output.name}"))
        except Exception as exc:
            tb = traceback.format_exc()
            self.after(0, lambda: self._log(tb))
            self.after(0, lambda: self._set_status("Failed — see log for details."))

    # ── Barchart scrape ────────────────────────────────────────────────────────

    def _run_barchart(self) -> None:
        try:
            url    = self.url_var.get().strip()
            output = Path(self.output_var.get().strip()).expanduser()

            self.after(0, lambda: self._log("Scraping Barchart futures table…"))
            df = scraper.scrape_barchart_futures(
                url,
                log=lambda m: self.after(0, lambda m=m: self._log(m)),
            )

            output.parent.mkdir(parents=True, exist_ok=True)
            with scraper.pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Barchart Futures")

            n = len(df)
            self.after(0, lambda: self._update_preview(df))
            self.after(0, lambda: self._log(f"✓ Saved {n} rows → {output.resolve()}"))
            self.after(0, lambda: self._set_status(f"Done — {n} rows saved to {output.name}"))
        except Exception as exc:
            tb = traceback.format_exc()
            self.after(0, lambda: self._log(tb))
            self.after(0, lambda: self._set_status("Failed — see log for details."))

    # ── Preview ────────────────────────────────────────────────────────────────

    def _update_preview(self, df) -> None:
        self._current_df = df
        cols = list(df.columns)
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = cols
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=max(100, min(220, len(col) * 10)),
                             stretch=True, anchor="w")
        for _, row in df.head(100).iterrows():
            self.tree.insert("", END, values=[str(v) for v in row.tolist()])
        self._log(f"Preview updated ({min(len(df), 100)} of {len(df)} rows shown).")


def _write_crash_log(text: str) -> None:
    """Write a crash log to the desktop so it's visible even when the app quits silently."""
    try:
        desktop = Path.home() / "Desktop" / "WSJ_Treasury_CRASH.txt"
        desktop.write_text(text, encoding="utf-8")
    except Exception:
        pass
    # Also try next to the executable
    try:
        log_path = HERE / "WSJ_Treasury_CRASH.txt"
        log_path.write_text(text, encoding="utf-8")
    except Exception:
        pass


def main() -> int:
    import traceback as _tb

    # Log startup info immediately so we know the app launched
    startup_info = (
        f"Python: {sys.version}\n"
        f"Executable: {sys.executable}\n"
        f"Frozen: {getattr(sys, 'frozen', False)}\n"
        f"HERE: {HERE}\n"
        f"sys.path: {sys.path}\n"
        f"scraper loaded: {scraper is not None}\n"
        f"import error: {_import_error}\n"
    )

    try:
        if scraper is None:
            msg = f"Could not import wsj_treasury_scraper_fixed:\n{_import_error}\n\n{startup_info}"
            _write_crash_log(msg)
            # Try to show a Tkinter error dialog before giving up
            try:
                import tkinter as _tk
                _root = _tk.Tk()
                _root.withdraw()
                from tkinter import messagebox as _mb
                _mb.showerror("Import Error",
                    f"Could not load scraper module:\n{_import_error}\n\nSee WSJ_Treasury_CRASH.txt on your Desktop.")
                _root.destroy()
            except Exception:
                pass
            return 1

        app = TreasuryApp()
        app.mainloop()
        return 0

    except Exception as exc:
        crash = f"CRASH:\n{_tb.format_exc()}\n\nStartup info:\n{startup_info}"
        _write_crash_log(crash)
        try:
            import tkinter as _tk
            _root = _tk.Tk()
            _root.withdraw()
            from tkinter import messagebox as _mb
            _mb.showerror("Crash", f"{exc}\n\nDetails saved to WSJ_Treasury_CRASH.txt on your Desktop.")
            _root.destroy()
        except Exception:
            pass
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
