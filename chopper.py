import os
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

EXCEL_MAX_ROWS = 1_048_576
DEFAULT_ROWS = EXCEL_MAX_ROWS - 10  # 1,048,566

# --- Optional drag & drop ---
DND_AVAILABLE = False
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False

# --- Optional Windows Excel COM for XLSB output ---
WIN32_AVAILABLE = False
if sys.platform.startswith("win"):
    try:
        import win32com.client as win32  # pip install pywin32
        WIN32_AVAILABLE = True
    except Exception:
        WIN32_AVAILABLE = False


class ChopperApp:
    def __init__(self, master):
        self.master = master
        self.master.title("File Chopper")
        self.master.minsize(720, 330)

        # State
        self.input_file = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.num_rows = tk.IntVar(value=DEFAULT_ROWS)
        self.status_text = tk.StringVar(value="Ready.")
        self.last_dir = os.path.expanduser("~")
        self.output_format = tk.StringVar(value="xlsx")  # csv | xlsx | xlsb

        # UI
        self._init_dark_theme()
        self._build_ui()

    # ---------------- UI ----------------
    def _init_dark_theme(self):
        style = ttk.Style(self.master)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        bg = "#121212"; fg = "#e5e5e5"; acc = "#2a2a2a"
        self.master.configure(bg=bg)
        style.configure(".", background=bg, foreground=fg, fieldbackground=acc)
        style.configure("TEntry", fieldbackground=acc)
        style.configure("TButton", background=acc, foreground=fg)
        style.map("TButton", background=[("active", "#3a3a3a")])
        style.configure("TLabel", background=bg, foreground=fg)
        style.configure("TRadiobutton", background=bg, foreground=fg)
        style.configure("TProgressbar", troughcolor=acc)

    def _build_ui(self):
        pad = {"padx": 10, "pady": 8}

        frm = ttk.Frame(self.master, padding=10)
        frm.grid(sticky="nsew")
        self.master.grid_columnconfigure(0, weight=1)
        frm.grid_columnconfigure(1, weight=1)

        # Input
        ttk.Label(frm, text="Input file (.csv/.xlsx):").grid(row=0, column=0, sticky="e", **pad)
        self.ent_input = ttk.Entry(frm, textvariable=self.input_file)
        self.ent_input.grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse…", command=self.browse_input_file).grid(row=0, column=2, **pad)

        hint = "Drag & drop enabled" if DND_AVAILABLE else "DnD not installed (pip install tkinterdnd2)"
        ttk.Label(frm, text=hint).grid(row=1, column=1, sticky="w", padx=10, pady=(0, 6))
        if DND_AVAILABLE:
            self.ent_input.drop_target_register(DND_FILES)
            self.ent_input.dnd_bind("<<Drop>>", self._on_drop_file)

        # Rows per file
        ttk.Label(frm, text="Rows per chunk:").grid(row=2, column=0, sticky="e", **pad)
        self.ent_rows = ttk.Entry(frm, textvariable=self.num_rows, width=16)
        self.ent_rows.grid(row=2, column=1, sticky="w", **pad)
        ttk.Label(frm, text="(default: 1,048,566)").grid(row=2, column=2, sticky="w", **pad)

        # Output dir
        ttk.Label(frm, text="Output folder:").grid(row=3, column=0, sticky="e", **pad)
        self.ent_out = ttk.Entry(frm, textvariable=self.output_dir)
        self.ent_out.grid(row=3, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Select…", command=self.browse_output_dir).grid(row=3, column=2, **pad)

        # Output format
        fmt_frame = ttk.Frame(frm)
        fmt_frame.grid(row=4, column=0, columnspan=3, sticky="w", padx=10, pady=(2, 2))
        ttk.Label(fmt_frame, text="Output format:").grid(row=0, column=0, sticky="w")
        for i, (text, val) in enumerate([("CSV", "csv"), ("XLSX", "xlsx"), ("XLSB", "xlsb")], start=1):
            ttk.Radiobutton(fmt_frame, text=text, value=val, variable=self.output_format).grid(row=0, column=i, padx=(12,0))

        # Progress + actions
        self.progress = ttk.Progressbar(frm, mode="determinate")
        self.progress.grid(row=5, column=0, columnspan=3, sticky="ew", padx=10, pady=(8, 2))

        self.lbl_status = ttk.Label(frm, textvariable=self.status_text)
        self.lbl_status.grid(row=6, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 6))

        self.btn_start = ttk.Button(frm, text="Start", command=self.start_chopping)
        self.btn_start.grid(row=7, column=1, sticky="e", padx=10, pady=10)
        ttk.Button(frm, text="Quit", command=self.master.destroy).grid(row=7, column=2, sticky="w", padx=10, pady=10)

    # ------------- Events -------------
    def _on_drop_file(self, event):
        raw = event.data.strip()
        if raw.startswith("{") and raw.endswith("}"):
            raw = raw[1:-1]
        path = raw.split("} {")[0] if "} {" in raw else raw
        path = path.strip()
        if os.path.isfile(path):
            self.input_file.set(path)
            self.last_dir = os.path.dirname(path)

    def browse_input_file(self):
        filename = filedialog.askopenfilename(
            title="Select input file",
            initialdir=self.last_dir,
            filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            self.last_dir = os.path.dirname(filename)

    def browse_output_dir(self):
        directory = filedialog.askdirectory(title="Select output folder", initialdir=self.last_dir)
        if directory:
            self.output_dir.set(directory)
            self.last_dir = directory

    # ------------- Core -------------
    def start_chopping(self):
        in_path = self.input_file.get().strip()
        out_dir = self.output_dir.get().strip()
        out_fmt = self.output_format.get()

        if not in_path:
            messagebox.showerror("Error", "Please select an input file.")
            return
        if not out_dir:
            messagebox.showerror("Error", "Please select an output folder.")
            return

        try:
            rows = int(self.num_rows.get())
            if rows <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror("Error", "Rows per chunk must be a positive integer.")
            return

        ext = os.path.splitext(in_path)[1].lower()
        if ext not in (".csv", ".xlsx"):
            messagebox.showerror("Error", "Unsupported input type. Use .csv or .xlsx.")
            return

        if out_fmt == "xlsb" and not (sys.platform.startswith("win") and WIN32_AVAILABLE):
            messagebox.showwarning(
                "XLSB not available",
                "XLSB export needs Excel + pywin32 on Windows. Falling back to XLSX."
            )
            self.output_format.set("xlsx")
            out_fmt = "xlsx"

        self._set_busy(True)
        t = threading.Thread(target=self._run_chop, args=(in_path, out_dir, rows, ext, out_fmt), daemon=True)
        t.start()

    def _set_busy(self, busy: bool):
        state = "disabled" if busy else "normal"
        for w in [self.btn_start, self.ent_input, self.ent_out, self.ent_rows]:
            w.configure(state=state)

    def _run_chop(self, in_path, out_dir, rows, in_ext, out_fmt):
        try:
            stem = os.path.splitext(os.path.basename(in_path))[0]
            if in_ext == ".csv":
                self._chop_csv_resilient(in_path, out_dir, stem, rows, out_fmt)
            else:
                self._chop_xlsx(in_path, out_dir, stem, rows, out_fmt)

            self._set_status("Done. Files saved to: " + out_dir)
            messagebox.showinfo("Success", "File has been chopped successfully.")
        except Exception as e:
            self._set_status("Failed.")
            messagebox.showerror("Error", f"An error occurred: {e}")
        finally:
            self._set_busy(False)
            self._progress_reset()

    # ------------- Encoding-robust CSV -------------
    def _csv_chunk_iter(self, path, rows):
        """
        Yield pandas chunks with progressively more lenient decoding.
        """
        encodings = ["utf-8", "utf-8-sig", "cp1252", "latin1"]
        last_err = None
        for enc in encodings:
            try:
                # Fast engine first; low_memory=False for stable dtypes
                return pd.read_csv(path, chunksize=rows, encoding=enc, low_memory=False)
            except Exception as e:
                last_err = e

        # Final fallback: python engine + skip bad lines
        try:
            return pd.read_csv(
                path, chunksize=rows, encoding="latin1",
                engine="python", on_bad_lines="skip", low_memory=False
            )
        except Exception as e:
            raise RuntimeError(
                f"Unable to decode CSV with common encodings. Last error: {last_err or e}"
            )

    def _chop_csv_resilient(self, in_path, out_dir, stem, rows, out_fmt):
        # Try to count rows in a binary-safe way; if it fails, go indeterminate
        total_rows = self._try_count_rows_binary(in_path)
        if total_rows is None:
            self._progress_indeterminate_start()
            self._set_status("Reading CSV (unknown encoding)…")
        else:
            chunks_est = max(1, (total_rows + rows - 1) // rows)
            self._set_status(f"CSV rows (approx): {total_rows:,}. Creating up to {chunks_est} chunk(s)…")
            self._set_progress_total(chunks_est)

        chunk_iter = self._csv_chunk_iter(in_path, rows)
        written = 0
        for idx, chunk in enumerate(chunk_iter, start=1):
            out_path = self._safe_out_path(out_dir, stem, idx, out_fmt)
            self._write_out(chunk, out_path, out_fmt)
            written += len(chunk)
            self._bump_progress(idx)  # determinate if we had a total; otherwise keep indeterminate
        if total_rows is None:
            self._progress_indeterminate_stop()

    def _try_count_rows_binary(self, path):
        """
        Count lines in binary to avoid decode errors, then subtract one for header if present.
        If the file is empty or has 0/1 lines, return 0; if unsure, return None.
        """
        try:
            with open(path, "rb") as f:
                # Count b'\n'; this can undercount last line if no newline—good enough for progress.
                data = f.read()
                lines = data.count(b"\n")
                if lines <= 0:
                    return 0
                # Heuristic: assume header exists; subtract 1
                return max(0, lines - 1)
        except Exception:
            return None

    # ------------- Excel input -------------
    def _chop_xlsx(self, in_path, out_dir, stem, rows, out_fmt):
        self._set_status("Loading Excel…")
        df = pd.read_excel(in_path, dtype=str)
        total = len(df)
        chunks = max(1, (total + rows - 1) // rows)
        self._set_status(f"Excel rows: {total:,}. Creating {chunks} chunk(s)…")
        self._set_progress_total(chunks)
        for idx, start in enumerate(range(0, total, rows), start=1):
            part = df.iloc[start:start + rows]
            out_path = self._safe_out_path(out_dir, stem, idx, out_fmt)
            self._write_out(part, out_path, out_fmt)
            self._bump_progress(idx)

    # ------------- Writers -------------
    def _write_out(self, df: pd.DataFrame, out_path: str, out_fmt: str):
        ext = os.path.splitext(out_path)[1].lower()
        stem = out_path[:-len(ext)]
        if out_fmt == "csv":
            # Excel-friendly UTF-8 with BOM
            df.to_csv(stem + ".csv", index=False, encoding="utf-8-sig")
        elif out_fmt == "xlsx":
            df.to_excel(stem + ".xlsx", index=False)
        elif out_fmt == "xlsb":
            # Write temp xlsx then convert with Excel COM
            temp_xlsx = stem + ".__tmp__.xlsx"
            df.to_excel(temp_xlsx, index=False)
            self._convert_xlsx_to_xlsb(temp_xlsx, stem + ".xlsb")
            try:
                os.remove(temp_xlsx)
            except Exception:
                pass

    def _convert_xlsx_to_xlsb(self, xlsx_path: str, xlsb_path: str):
        if not (sys.platform.startswith("win") and WIN32_AVAILABLE):
            raise RuntimeError("XLSB conversion requires Windows + Excel (pywin32).")
        self._set_status("Converting to XLSB via Excel…")
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        try:
            wb = excel.Workbooks.Open(xlsx_path)
            # 50 = xlExcel12 (XLSB)
            wb.SaveAs(xlsb_path, FileFormat=50)
            wb.Close(SaveChanges=False)
        finally:
            excel.Quit()

    # ------------- Progress helpers -------------
    def _set_progress_total(self, total):
        self.progress.configure(mode="determinate", maximum=total, value=0)
        self.master.update_idletasks()

    def _bump_progress(self, current):
        if self.progress["mode"] == "determinate":
            self.progress["value"] = current
        self.master.update_idletasks()

    def _progress_indeterminate_start(self):
        self.progress.configure(mode="indeterminate")
        self.progress.start(60)

    def _progress_indeterminate_stop(self):
        self.progress.stop()
        self.progress.configure(mode="determinate", value=0)

    def _progress_reset(self):
        try:
            self.progress.stop()
        except Exception:
            pass
        self.progress.configure(mode="determinate", value=0)

    # ------------- Misc -------------
    def _safe_out_path(self, out_dir, stem, idx, out_fmt):
        ext = "." + out_fmt
        candidate = os.path.join(out_dir, f"{stem}_{idx}{ext}")
        n = 1
        while os.path.exists(candidate):
            candidate = os.path.join(out_dir, f"{stem}_{idx}({n}){ext}")
            n += 1
        return candidate

    def _set_status(self, text):
        self.status_text.set(text)
        self.master.update_idletasks()


def main():
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    app = ChopperApp(root)
    root.update_idletasks()
    w, h = 820, 360
    x = (root.winfo_screenwidth() // 2) - (w // 2)
    y = (root.winfo_screenheight() // 2) - (h // 2)
    root.geometry(f"{w}x{h}+{x}+{y}")
    root.mainloop()


if __name__ == "__main__":
    main()
