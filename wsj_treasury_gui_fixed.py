#!/usr/bin/env python3
"""Desktop GUI for the WSJ Treasury / Barchart Futures scraper."""

from __future__ import annotations

import sys
import threading
import traceback
from pathlib import Path
from tkinter import (
    BOTH, END, LEFT, RIGHT, X, Y, W,
    BooleanVar, Button, Entry, Frame, Label, StringVar, Text, Tk,
    filedialog, messagebox,
)
from tkinter import ttk

# ── Path resolution ────────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    HERE = Path(sys.executable).resolve().parent
    _bundle = Path(getattr(sys, "_MEIPASS", HERE))
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

# matplotlib — optional, chart tab is disabled if not installed
try:
    import matplotlib
    matplotlib.use("TkAgg")
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
    _MPL = True
except ImportError:
    _MPL = False

WSJ_URL      = "https://www.wsj.com/market-data/bonds/treasuries#treasuryB"
BARCHART_URL = "https://www.barchart.com/futures/financials?viewName=main"
SOURCE_WSJ      = "WSJ Treasuries"
SOURCE_BARCHART = "Barchart Futures"
SOURCES = [SOURCE_WSJ, SOURCE_BARCHART]

# Columns to skip as chart series (used as X axis or non-numeric)
_SKIP_COLS = {"Maturity", "Article Date", "Last Payment Date", "Symbol",
              "Contract Name", "Time"}


class TreasuryApp(Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Treasury / Futures Scraper")
        self.geometry("1280x860")
        self.minsize(1000, 700)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.source_var    = StringVar(value=SOURCE_WSJ)
        self.url_var       = StringVar(value=WSJ_URL)
        self.output_var    = StringVar(value=str(HERE / "wsj_treasuries.xlsx"))
        self.reference_var = StringVar(value="")
        self.status_var    = StringVar(value="Ready.")

        self._worker: threading.Thread | None = None
        self._current_df = None
        self._chart_vars: dict[str, BooleanVar] = {}   # column → checkbox var

        self._build_ui()
        if _import_error is not None:
            self._log(f"Import warning: {_import_error}")
            self._set_status("The scraper module could not be imported.")

    # ── UI ─────────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        outer = Frame(self, padx=12, pady=12)
        outer.pack(fill=BOTH, expand=True)

        # ── Controls ───────────────────────────────────────────────────────────
        top = Frame(outer)
        top.pack(fill=X)

        Label(top, text="Data source").grid(row=0, column=0, sticky="w", pady=(0, 4))
        src_frame = Frame(top)
        src_frame.grid(row=0, column=1, sticky="w", padx=(8, 8), pady=(0, 4))
        for src in SOURCES:
            ttk.Radiobutton(src_frame, text=src, variable=self.source_var,
                            value=src, command=self._on_source_changed).pack(side=LEFT, padx=(0, 16))

        Label(top, text="URL").grid(row=1, column=0, sticky="w", pady=(0, 4))
        Entry(top, textvariable=self.url_var).grid(row=1, column=1, sticky="we", padx=(8, 8), pady=(0, 4))

        Label(top, text="Output file").grid(row=2, column=0, sticky="w", pady=(0, 4))
        out_row = Frame(top)
        out_row.grid(row=2, column=1, sticky="we", padx=(8, 8), pady=(0, 4))
        Entry(out_row, textvariable=self.output_var).pack(side=LEFT, fill=X, expand=True)
        Button(out_row, text="Browse…", command=self._browse_output).pack(side=RIGHT, padx=(8, 0))

        self._ref_label = Label(top, text="Reference date (YYYY-MM-DD)")
        self._ref_label.grid(row=3, column=0, sticky="w", pady=(0, 4))
        self._ref_entry = Entry(top, textvariable=self.reference_var)
        self._ref_entry.grid(row=3, column=1, sticky="we", padx=(8, 8), pady=(0, 4))

        btn_row = Frame(top)
        btn_row.grid(row=4, column=1, sticky="w", padx=(8, 8), pady=(6, 2))
        Button(btn_row, text="Run scrape",         command=self._run_clicked,        width=16).pack(side=LEFT)
        Button(btn_row, text="Open output folder", command=self._open_output_folder, width=18).pack(side=LEFT, padx=(8, 0))

        log_btn_row = Frame(top)
        log_btn_row.grid(row=5, column=1, sticky="w", padx=(8, 8), pady=(2, 8))
        Button(log_btn_row, text="Save log…", command=self._save_log,  width=14).pack(side=LEFT)
        Button(log_btn_row, text="Clear log", command=self._clear_log, width=12).pack(side=LEFT, padx=(8, 0))

        top.columnconfigure(1, weight=1)

        # ── Middle: notebook + log ─────────────────────────────────────────────
        mid = Frame(outer)
        mid.pack(fill=BOTH, expand=True, pady=(8, 8))

        # Notebook (Preview + Chart tabs)
        self.notebook = ttk.Notebook(mid)
        self.notebook.pack(side=LEFT, fill=BOTH, expand=True)

        # ── Tab 1: Preview ─────────────────────────────────────────────────────
        preview_tab = Frame(self.notebook)
        self.notebook.add(preview_tab, text="Preview")

        yscroll = ttk.Scrollbar(preview_tab, orient="vertical")
        xscroll = ttk.Scrollbar(preview_tab, orient="horizontal")
        yscroll.pack(side=RIGHT, fill=Y)
        xscroll.pack(side="bottom", fill=X)

        self.tree = ttk.Treeview(preview_tab, show="headings", height=18,
                                 yscrollcommand=yscroll.set,
                                 xscrollcommand=xscroll.set)
        yscroll.config(command=self.tree.yview)
        xscroll.config(command=self.tree.xview)
        self.tree.pack(fill=BOTH, expand=True)

        # ── Tab 2: Chart ───────────────────────────────────────────────────────
        chart_tab = Frame(self.notebook)
        self.notebook.add(chart_tab, text="Chart (WSJ)")

        if _MPL:
            # Left: checkboxes to pick series
            ctrl_frame = Frame(chart_tab, width=180)
            ctrl_frame.pack(side=LEFT, fill=Y, padx=(8, 4), pady=8)
            ctrl_frame.pack_propagate(False)

            Label(ctrl_frame, text="Series", font=("", 10, "bold")).pack(anchor=W, pady=(0, 4))
            self._checkbox_frame = Frame(ctrl_frame)
            self._checkbox_frame.pack(fill=X, anchor=W)

            Label(ctrl_frame, text="X axis", font=("", 10, "bold")).pack(anchor=W, pady=(10, 4))
            self._x_var = StringVar(value="Maturity")
            self._x_menu = ttk.Combobox(ctrl_frame, textvariable=self._x_var,
                                        state="readonly", width=18)
            self._x_menu.pack(anchor=W)
            self._x_menu.bind("<<ComboboxSelected>>", lambda _: self._draw_chart())

            Button(ctrl_frame, text="Update chart", command=self._draw_chart,
                   width=16).pack(anchor=W, pady=(12, 0))

            # Right: matplotlib canvas
            canvas_frame = Frame(chart_tab)
            canvas_frame.pack(side=LEFT, fill=BOTH, expand=True, pady=8, padx=(0, 8))

            self._fig = Figure(figsize=(7, 4), dpi=96, tight_layout=True)
            self._ax  = self._fig.add_subplot(111)
            self._canvas = FigureCanvasTkAgg(self._fig, master=canvas_frame)
            self._canvas.get_tk_widget().pack(fill=BOTH, expand=True)
            NavigationToolbar2Tk(self._canvas, canvas_frame).pack(fill=X)

            self._ax.set_title("Run a WSJ scrape to populate the chart")
            self._ax.set_xlabel("Maturity")
            self._canvas.draw()
        else:
            Label(chart_tab,
                  text="matplotlib is not installed.\nRun:  pip install matplotlib",
                  justify="center", pady=40).pack(expand=True)

        # ── Log panel ──────────────────────────────────────────────────────────
        log_box = Frame(mid, width=380)
        log_box.pack(side=RIGHT, fill=Y, padx=(8, 0))
        log_box.pack_propagate(False)
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
            if not messagebox.askyesno("Scrape running", "A scrape is still running. Quit anyway?"):
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
            title="Save log", defaultextension=".txt",
            filetypes=[("Text file", "*.txt"), ("All files", "*.*")],
            initialfile="scraper_log.txt", initialdir=str(HERE),
        )
        if path:
            try:
                Path(path).write_text(self.log.get("1.0", END), encoding="utf-8")
                self._log(f"Log saved to {path}")
            except Exception as exc:
                messagebox.showerror("Save failed", str(exc))

    def _browse_output(self) -> None:
        filename = filedialog.asksaveasfilename(
            title="Choose output Excel file", defaultextension=".xlsx",
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
            self.after(0, lambda: self._update_preview(df, is_wsj=True))
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
                url, log=lambda m: self.after(0, lambda m=m: self._log(m)),
            )

            output.parent.mkdir(parents=True, exist_ok=True)
            with scraper.pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Barchart Futures")

            n = len(df)
            self.after(0, lambda: self._update_preview(df, is_wsj=False))
            self.after(0, lambda: self._log(f"✓ Saved {n} rows → {output.resolve()}"))
            self.after(0, lambda: self._set_status(f"Done — {n} rows saved to {output.name}"))
        except Exception as exc:
            tb = traceback.format_exc()
            self.after(0, lambda: self._log(tb))
            self.after(0, lambda: self._set_status("Failed — see log for details."))

    # ── Preview & Chart ────────────────────────────────────────────────────────

    def _update_preview(self, df, is_wsj: bool = False) -> None:
        self._current_df = df
        cols = list(df.columns)

        # Update treeview
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = cols
        COL_WIDTH = 130
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=COL_WIDTH, minwidth=COL_WIDTH,
                             stretch=False, anchor="w")
        for _, row in df.head(100).iterrows():
            self.tree.insert("", END, values=[str(v) for v in row.tolist()])
        self._log(f"Preview updated ({min(len(df), 100)} of {len(df)} rows shown).")

        # Update chart controls (WSJ only)
        if _MPL and is_wsj:
            self._rebuild_chart_controls(df)
            self._draw_chart()

    def _rebuild_chart_controls(self, df) -> None:
        """Rebuild the series checkboxes and X-axis dropdown from the new dataframe."""
        # Clear old checkboxes
        for widget in self._checkbox_frame.winfo_children():
            widget.destroy()
        self._chart_vars.clear()

        # Numeric columns only (exclude known non-numeric / date cols)
        import pandas as pd
        numeric_cols = [
            c for c in df.columns
            if c not in _SKIP_COLS and pd.api.types.is_numeric_dtype(
                pd.to_numeric(df[c], errors="coerce")
            )
        ]

        for col in numeric_cols:
            var = BooleanVar(value=True)
            self._chart_vars[col] = var
            ttk.Checkbutton(self._checkbox_frame, text=col, variable=var,
                            command=self._draw_chart).pack(anchor=W)

        # X axis dropdown — date/text columns
        x_options = [c for c in df.columns if c not in numeric_cols]
        self._x_menu["values"] = x_options
        if "Maturity" in x_options:
            self._x_var.set("Maturity")
        elif x_options:
            self._x_var.set(x_options[0])

    def _draw_chart(self) -> None:
        """Redraw the matplotlib chart based on current checkbox selections."""
        if not _MPL or self._current_df is None:
            return

        import pandas as pd
        import matplotlib.dates as mdates
        import matplotlib.ticker as mticker

        df = self._current_df.copy()
        x_col = self._x_var.get()
        selected = [col for col, var in self._chart_vars.items() if var.get()]

        self._ax.clear()

        if not selected:
            self._ax.set_title("No series selected")
            self._canvas.draw()
            return

        if x_col not in df.columns:
            self._ax.set_title(f"X column '{x_col}' not found")
            self._canvas.draw()
            return

        # Try to parse X as dates so we get a proper time axis
        x_as_dates = pd.to_datetime(df[x_col], errors="coerce")
        use_dates = x_as_dates.notna().sum() > len(df) * 0.5

        if use_dates:
            df["_x"] = x_as_dates
        else:
            df["_x"] = range(len(df))

        # Sort by X so the line runs left→right
        df = df.sort_values("_x").reset_index(drop=True)

        for col in selected:
            y_vals = pd.to_numeric(df[col], errors="coerce")
            self._ax.plot(df["_x"], y_vals, marker="o", markersize=3,
                         linewidth=1.2, label=col)

        self._ax.set_xlabel(x_col)
        self._ax.set_title("WSJ Treasury Data")
        self._ax.legend(fontsize=8, loc="best")
        self._ax.grid(True, linestyle="--", alpha=0.4)

        if use_dates:
            # Auto-format: pick a sensible locator based on date range
            date_range = (df["_x"].max() - df["_x"].min()).days
            if date_range > 365 * 5:
                self._ax.xaxis.set_major_locator(mdates.YearLocator(2))
                self._ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
            elif date_range > 365:
                self._ax.xaxis.set_major_locator(mdates.YearLocator())
                self._ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
            else:
                self._ax.xaxis.set_major_locator(mdates.MonthLocator())
                self._ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
        else:
            # Numeric index: show at most 20 evenly-spaced ticks
            n = len(df)
            step = max(1, n // 20)
            ticks = list(range(0, n, step))
            labels = [str(df[x_col].iloc[i]) for i in ticks]
            self._ax.set_xticks(ticks)
            self._ax.set_xticklabels(labels)

        self._fig.autofmt_xdate(rotation=45, ha="right")
        self._canvas.draw()
        # Switch to chart tab automatically
        self.notebook.select(1)


def _write_crash_log(text: str) -> None:
    for path in [Path.home() / "Desktop" / "WSJ_Treasury_CRASH.txt",
                 HERE / "WSJ_Treasury_CRASH.txt"]:
        try:
            path.write_text(text, encoding="utf-8")
        except Exception:
            pass


def main() -> int:
    import traceback as _tb
    startup_info = (
        f"Python: {sys.version}\n"
        f"Executable: {sys.executable}\n"
        f"Frozen: {getattr(sys, 'frozen', False)}\n"
        f"HERE: {HERE}\n"
        f"scraper loaded: {scraper is not None}\n"
        f"import error: {_import_error}\n"
        f"matplotlib: {_MPL}\n"
    )
    try:
        if scraper is None:
            _write_crash_log(f"Could not import scraper:\n{_import_error}\n\n{startup_info}")
            try:
                import tkinter as _tk; _root = _tk.Tk(); _root.withdraw()
                from tkinter import messagebox as _mb
                _mb.showerror("Import Error", f"Could not load scraper:\n{_import_error}")
                _root.destroy()
            except Exception:
                pass
            return 1
        app = TreasuryApp()
        app.mainloop()
        return 0
    except Exception as exc:
        _write_crash_log(f"CRASH:\n{_tb.format_exc()}\n\n{startup_info}")
        try:
            import tkinter as _tk; _root = _tk.Tk(); _root.withdraw()
            from tkinter import messagebox as _mb
            _mb.showerror("Crash", f"{exc}\n\nSee WSJ_Treasury_CRASH.txt on Desktop.")
            _root.destroy()
        except Exception:
            pass
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
