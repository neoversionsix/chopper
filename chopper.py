import os
import sys
import io
import csv
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
    from tkinterdnd2 import DND_FILES, TkinterDnD  # pip install tkinterdnd2
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

# --- Optional charset detection ---
CHARDET_AVAILABLE = False
try:
    from charset_normalizer import from_bytes  # pip install charset-normalizer
    CHARDET_AVAILABLE = True
except Exception:
    CHARDET_AVAILABLE = False


class ChopperApp:
    def __init__(self, master):
        self.master = master
        self.master.title("File Chopper")
        self.master.minsize(820, 360)

        # State
        self.input_file = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.num_rows = tk.IntVar(value=DEFAULT_ROWS)
        self.status_text = tk.StringVar(value="Ready.")
        self.last_dir = os.path.expanduser("~")
        self.output_format = tk.StringVar(value="xlsx")  # csv | xlsx | xlsb
        self.temp_paths = []  # cleaned temp files to delete on exit

        # UI
        self._init_dark_theme()
        self._build_ui()
        self.master.protocol("WM_DELETE_WINDOW", self._on_close)

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
        frm = ttk.Frame(self.master, padding=10); frm.grid(sticky="nsew")
        self.master.grid_columnconfigure(0, weight=1); frm.grid_columnconfigure(1, weight=1)

        ttk.Label(frm, text="Input file (.csv/.xlsx):").grid(row=0, column=0, sticky="e", **pad)
        self.ent_input = ttk.Entry(frm, textvariable=self.input_file)
        self.ent_input.grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse…", command=self.browse_input_file).grid(row=0, column=2, **pad)

        hint = "Drag & drop enabled" if DND_AVAILABLE else "DnD not installed (pip install tkinterdnd2)"
        ttk.Label(frm, text=hint).grid(row=1, column=1, sticky="w", padx=10, pady=(0, 6))
        if DND_AVAILABLE:
            self.ent_input.drop_target_register(DND_FILES)
            self.ent_input.dnd_bind("<<Drop>>", self._on_drop_file)

        ttk.Label(frm, text="Rows per chunk:").grid(row=2, column=0, sticky="e", **pad)
        self.ent_rows = ttk.Entry(frm, textvariable=self.num_rows, width=16)
        self.ent_rows.grid(row=2, column=1, sticky="w", **pad)
        ttk.Label(frm, text="(default: 1,048,566)").grid(row=2, column=2, sticky="w", **pad)

        ttk.Label(frm, text="Output folder:").grid(row=3, column=0, sticky="e", **pad)
        self.ent_out = ttk.Entry(frm, textvariable=self.output_dir)
        self.ent_out.grid(row=3, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Select…", command=self.browse_output_dir).grid(row=3, column=2, **pad)

        fmt_frame = ttk.Frame(frm); fmt_frame.grid(row=4, column=0, columnspan=3, sticky="w", padx=10, pady=(2,2))
        ttk.Label(fmt_frame, text="Output format:").grid(row=0, column=0, sticky="w")
        for i, (text, val) in enumerate([("CSV", "csv"), ("XLSX", "xlsx"), ("XLSB", "xlsb")], start=1):
            ttk.Radiobutton(fmt_frame, text=text, value=val, variable=self.output_format).grid(row=0, column=i, padx=(12,0))

        self.progress = ttk.Progressbar(frm, mode="determinate")
        self.progress.grid(row=5, column=0, columnspan=3, sticky="ew", padx=10, pady=(8,2))
        self.lbl_status = ttk.Label(frm, textvariable=self.status_text)
        self.lbl_status.grid(row=6, column=0, columnspan=3, sticky="w", padx=10, pady=(0,6))

        self.btn_start = ttk.Button(frm, text="Start", command=self.start_chopping)
        self.btn_start.grid(row=7, column=1, sticky="e", padx=10, pady=10)
        ttk.Button(frm, text="Quit", command=self._on_close).grid(row=7, column=2, sticky="w", padx=10, pady=10)

    # ------------- Events -------------
    def _on_drop_file(self, event):
        raw = event.data.strip()
        if raw.startswith("{") and raw.endswith("}"): raw = raw[1:-1]
        path = raw.split("} {")[0] if "} {" in raw else raw
        if os.path.isfile(path):
            self.input_file.set(path)
            self.last_dir = os.path.dirname(path)

    def browse_input_file(self):
        fn = filedialog.askopenfilename(
            title="Select input file",
            initialdir=self.last_dir,
            filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx"), ("All files", "*.*")]
        )
        if fn:
            self.input_file.set(fn)
            self.last_dir = os.path.dirname(fn)

    def browse_output_dir(self):
        d = filedialog.askdirectory(title="Select output folder", initialdir=self.last_dir)
        if d:
            self.output_dir.set(d); self.last_dir = d

    # ------------- Core -------------
    def start_chopping(self):
        in_path = self.input_file.get().strip()
        out_dir = self.output_dir.get().strip()
        out_fmt = self.output_format.get()

        if not in_path: return messagebox.showerror("Error", "Please select an input file.")
        if not out_dir: return messagebox.showerror("Error", "Please select an output folder.")

        try:
            rows = int(self.num_rows.get())
            if rows <= 0: raise ValueError
        except Exception:
            return messagebox.showerror("Error", "Rows per chunk must be a positive integer.")

        ext = os.path.splitext(in_path)[1].lower()
        if ext not in (".csv", ".xlsx"):
            return messagebox.showerror("Error", "Unsupported input type. Use .csv or .xlsx.")

        if out_fmt == "xlsb" and not (sys.platform.startswith("win") and WIN32_AVAILABLE):
            messagebox.showwarning("XLSB not available", "XLSB export needs Excel + pywin32 on Windows. Falling back to XLSX.")
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
                total_rows_bin = self._count_lines_binary(in_path)
                if total_rows_bin is None:
                    self._progress_indeterminate_start()
                else:
                    approx = max(0, total_rows_bin - 1)
                    est_chunks = max(1, (approx + rows - 1) // rows)
                    self._set_progress_total(est_chunks)

                clean_path, enc_used = self._transcode_csv_to_utf8_temp(in_path)
                delim = self._sniff_delimiter(clean_path) or ","
                self._set_status(f"Cleaned CSV → UTF-8 ({enc_used}), delimiter='{delim}'.")
                self._chop_clean_csv(clean_path, out_dir, stem, rows, out_fmt, delimiter=delim)

                if total_rows_bin is None:
                    self._progress_indeterminate_stop()

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

    # ------------- Cleaning / detection -------------
    def _detect_encoding(self, sample_bytes: bytes) -> str:
        if CHARDET_AVAILABLE:
            res = from_bytes(sample_bytes).best()
            if res and res.encoding:
                return res.encoding
        for enc in ("utf-8", "utf-8-sig", "cp1252", "latin1"):
            try:
                sample_bytes.decode(enc)
                return enc
            except Exception:
                continue
        return "latin1"

    def _transcode_csv_to_utf8_temp(self, path: str):
        with open(path, "rb") as f:
            sample = f.read(128 * 1024)
        enc = self._detect_encoding(sample)

        tmp_fd, tmp_path = tempfile.mkstemp(prefix="chopper_clean_", suffix=".csv")
        os.close(tmp_fd)
        self.temp_paths.append(tmp_path)

        self._set_status(f"Cleaning input (encoding≈{enc})…")
        with open(path, "rb") as fin, io.open(tmp_path, "w", encoding="utf-8", newline="\n") as fout:
            decoder = io.TextIOWrapper(fin, encoding=enc, errors="replace", newline=None)
            while True:
                chunk = decoder.read(1_000_000)
                if not chunk:
                    break
                chunk = chunk.replace("\ufeff", "").replace("\xa0", " ").replace("\x00", "")
                fout.write(chunk)

        return tmp_path, enc

    def _sniff_delimiter(self, utf8_path: str):
        try:
            with open(utf8_path, "r", encoding="utf-8", newline="") as f:
                sample = f.read(64 * 1024)
            dialect = csv.Sniffer().sniff(sample, delimiters=[",",";","\t","|"])
            return dialect.delimiter
        except Exception:
            return None

    # ------------- CSV chopping (using cleaned UTF-8) -------------
    def _chop_clean_csv(self, clean_path, out_dir, stem, rows, out_fmt, delimiter=","):
        i = 0
        for i, chunk in enumerate(
            pd.read_csv(clean_path, chunksize=rows, encoding="utf-8", low_memory=False, sep=delimiter),
            start=1
        ):
            out_path = self._safe_out_path(out_dir, stem, i, out_fmt)
            self._write_out(chunk, out_path, out_fmt)
            self._bump_progress(i)
        if i == 0:
            self._set_status("No data rows found after cleaning.")

    # ------------- Helpers -------------
    def _count_lines_binary(self, path):
        try:
            with open(path, "rb") as f:
                buf = f.read()
                return buf.count(b"\n") or buf.count(b"\r")
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
            df.to_csv(stem + ".csv", index=False, encoding="utf-8-sig")
        elif out_fmt == "xlsx":
            df.to_excel(stem + ".xlsx", index=False)
        elif out_fmt == "xlsb":
            temp_xlsx = stem + ".__tmp__.xlsx"
            df.to_excel(temp_xlsx, index=False)
            try:
                self._convert_xlsx_to_xlsb(temp_xlsx, stem + ".xlsb")
            except Exception as e:
                # Fallback to XLSX if XLSB fails for any reason
                fallback = stem + ".xlsx"
                try:
                    if not os.path.exists(fallback):
                        os.replace(temp_xlsx, fallback)
                    else:
                        i = 1
                        fb = f"{stem} (fallback {i}).xlsx"
                        while os.path.exists(fb):
                            i += 1
                            fb = f"{stem} (fallback {i}).xlsx"
                        os.replace(temp_xlsx, fb)
                except Exception:
                    pass
                raise RuntimeError(f"XLSB conversion failed; wrote XLSX instead. Details: {e}")
            finally:
                try:
                    if os.path.exists(temp_xlsx):
                        os.remove(temp_xlsx)
                except Exception:
                    pass

    # --------- Hardened XLSB conversion (Excel COM) ----------
    def _abs_long_path(self, p: str) -> str:
        p = os.path.abspath(p)
        if sys.platform.startswith("win"):
            if not p.startswith("\\\\?\\"):
                p = "\\\\?\\" + p.replace("/", "\\")
        return p

    def _close_workbook_if_open(self, excel, target_fullname: str):
        try:
            for wb in list(excel.Workbooks):
                try:
                    if os.path.normcase(str(wb.FullName)) == os.path.normcase(target_fullname):
                        wb.Close(SaveChanges=False)
                except Exception:
                    continue
        except Exception:
            pass

    def _convert_xlsx_to_xlsb(self, xlsx_path: str, xlsb_path: str):
        if not (sys.platform.startswith("win") and WIN32_AVAILABLE):
            raise RuntimeError("XLSB conversion requires Windows + Excel (pywin32).")

        os.makedirs(os.path.dirname(xlsb_path) or ".", exist_ok=True)

        xlsx_abs = self._abs_long_path(xlsx_path)
        xlsb_abs = self._abs_long_path(xlsb_path)

        try:
            if os.path.exists(xlsb_path):
                os.remove(xlsb_path)
        except Exception:
            base, ext = os.path.splitext(xlsb_path)
            i = 1
            new_target = f"{base} ({i}){ext}"
            while os.path.exists(new_target):
                i += 1
                new_target = f"{base} ({i}){ext}"
            xlsb_path = new_target
            xlsb_abs = self._abs_long_path(xlsb_path)

        excel = win32.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False

        last_err = None
        try:
            for attempt in range(1, 4):
                try:
                    wb = excel.Workbooks.Open(xlsx_abs)
                    try:
                        self._close_workbook_if_open(excel, xlsb_abs)
                        wb.CheckCompatibility = False
                        # 50 = xlExcel12 (.xlsb)
                        wb.SaveAs(xlsb_abs, FileFormat=50, Local=True)
                    finally:
                        wb.Close(SaveChanges=False)

                    if os.path.exists(xlsb_path) and os.path.getsize(xlsb_path) > 0:
                        return
                    raise RuntimeError("Excel reported success but target file missing/empty.")
                except Exception as e:
                    last_err = e
                    # backoff + tweak filename to dodge locks/collisions
                    import time, random
                    time.sleep(0.6 * attempt)
                    base, ext = os.path.splitext(xlsb_path)
                    xlsb_path = f"{base} (retry {attempt}){ext}"
                    xlsb_abs = self._abs_long_path(xlsb_path)
            raise RuntimeError(f"Excel SaveAs to XLSB failed after retries: {last_err}")
        finally:
            try:
                excel.Quit()
            except Exception:
                pass

    # ------------- Progress helpers -------------
    def _set_progress_total(self, total):
        self.progress.configure(mode="determinate", maximum=total, value=0)
        self.master.update_idletasks()

    def _bump_progress(self, current):
        if self.progress["mode"] == "determinate":
            self.progress["value"] = current
        self.master.update_idletasks()

    def _progress_indeterminate_start(self):
        self.progress.configure(mode="indeterminate"); self.progress.start(60)

    def _progress_indeterminate_stop(self):
        try: self.progress.stop()
        except Exception: pass
        self.progress.configure(mode="determinate", value=0)

    def _progress_reset(self):
        try: self.progress.stop()
        except Exception: pass
        self.progress.configure(mode="determinate", value=0)

    # ------------- Misc -------------
    def _safe_out_path(self, out_dir, stem, idx, out_fmt):
        ext = "." + out_fmt
        candidate = os.path.join(out_dir, f"{stem}_{idx}{ext}")
        n = 1
        while os.path.exists(candidate):
            candidate = os.path.join(out_dir, f"{stem}_{idx}({n}){ext}"); n += 1
        return candidate

    def _set_status(self, text):
        self.status_text.set(text); self.master.update_idletasks()

    def _on_close(self):
        for p in self.temp_paths:
            try: os.remove(p)
            except Exception: pass
        self.master.destroy()


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
