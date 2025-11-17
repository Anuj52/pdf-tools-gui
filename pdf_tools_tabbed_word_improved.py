#!/usr/bin/env python3
"""
pdf_tools_tabbed_word_improved.py

Single-file GUI-only app with top tabs:
 - Tab 1: Password Tool (bulk decrypt / re-encrypt)
 - Tab 2: Merge PDFs (reorder + merge, Add Folder included)
 - Tab 3: Word -> PDF (uses docx2pdf; requires Microsoft Word on Windows)

Enhancements:
 - Word->PDF tab shows a result table: File | Status | Message | Output Path
 - "Open Output Folder" button
 - "Preserve Subfolders" option when adding folders for Word→PDF (keeps relative structure)
 - README text and a PyInstaller .spec example included below (not written to disk automatically)

Usage:
    python pdf_tools_tabbed_word_improved.py

Dependencies:
    pip install PyPDF2 pandas docx2pdf tkinterdnd2
    (tkinterdnd2 optional for drag & drop; pandas optional for CSV parsing)
"""

import os
import shutil
import csv
import threading
import traceback
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from queue import Queue, Empty
import sys
import subprocess

from PyPDF2 import PdfReader, PdfWriter

# Optional libs
try:
    import pandas as pd
except Exception:
    pd = None

# docx2pdf (Windows + Word required)
try:
    from docx2pdf import convert as docx2pdf_convert
    docx2pdf_available = True
except Exception:
    docx2pdf_available = False

# GUI
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Optional drag & drop
tkdnd_available = False
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    tkdnd_available = True
except Exception:
    tkdnd_available = False

# -------------------------
# Utilities
# -------------------------
def timestamp_str():
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def safe_copy(src, dst):
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(src, dst)

def is_decrypt_success(result):
    return (result is True) or (isinstance(result, int) and result != 0)

def open_folder_in_explorer(path):
    path = os.path.abspath(path)
    if not os.path.exists(path):
        return False
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform.startswith("darwin"):
            subprocess.run(["open", path])
        else:
            subprocess.run(["xdg-open", path])
        return True
    except Exception:
        return False

# -------------------------
# Password CSV loader
# -------------------------
def load_password_csv(csv_path):
    mapping = {}
    try:
        if pd:
            df = pd.read_csv(csv_path, dtype=str)
            if 'filename' in df.columns and 'password' in df.columns:
                for _, row in df.iterrows():
                    mapping[str(row['filename']).strip()] = str(row['password'])
            else:
                for _, row in df.iterrows():
                    mapping[str(row.iloc[0]).strip()] = str(row.iloc[1])
        else:
            with open(csv_path, newline='', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                for row in reader:
                    if not row:
                        continue
                    if len(row) >= 2:
                        mapping[row[0].strip()] = row[1].strip()
        return mapping
    except Exception as e:
        raise RuntimeError(f"Failed to read CSV mapping: {e}")

# -------------------------
# Password processing core
# -------------------------
def process_single_pdf(
    file_path,
    common_password=None,
    per_file_map=None,
    new_password=None,
    output_folder=None,
    overwrite=False,
    backup_root=None,
    skip_unlocked=True,
):
    filename = os.path.basename(file_path)
    try:
        with open(file_path, "rb") as f:
            reader = PdfReader(f)

            if not reader.is_encrypted:
                if skip_unlocked and not new_password:
                    return filename, True, "Already unlocked (skipped)"
                writer = PdfWriter()
                for p in reader.pages:
                    writer.add_page(p)

                if output_folder and not overwrite:
                    os.makedirs(output_folder, exist_ok=True)
                    out_path = os.path.join(output_folder, filename)
                else:
                    out_path = file_path

                if overwrite and backup_root:
                    os.makedirs(backup_root, exist_ok=True)
                    safe_copy(file_path, os.path.join(backup_root, filename))

                with open(out_path, "wb") as out_f:
                    if new_password:
                        writer.encrypt(new_password)
                    writer.write(out_f)

                return filename, True, "Unlocked → Encrypted with new password" if new_password else "Already unlocked (rewritten)"

            per_pw = None
            if per_file_map and filename in per_file_map:
                per_pw = per_file_map[filename]

            success = False
            msg = "Failed to decrypt"
            for pw in (per_pw, common_password):
                if not pw:
                    continue
                try:
                    result = reader.decrypt(pw)
                except Exception:
                    result = 0
                if is_decrypt_success(result):
                    writer = PdfWriter()
                    for p in reader.pages:
                        writer.add_page(p)

                    if output_folder and not overwrite:
                        os.makedirs(output_folder, exist_ok=True)
                        out_path = os.path.join(output_folder, filename)
                    else:
                        out_path = file_path

                    if overwrite and backup_root:
                        os.makedirs(backup_root, exist_ok=True)
                        safe_copy(file_path, os.path.join(backup_root, filename))

                    with open(out_path, "wb") as out_f:
                        if new_password:
                            writer.encrypt(new_password)
                        writer.write(out_f)

                    used = 'per-file' if (pw == per_pw and per_pw) else 'common'
                    success = True
                    msg = f"Decrypted (used password: {used})"
                    break
                else:
                    msg = "Incorrect password"
            return filename, success, msg
    except Exception as e:
        return filename, False, f"Error: {e}"

def run_batch(
    file_list,
    common_password=None,
    per_file_map=None,
    new_password=None,
    output_folder=None,
    overwrite=False,
    backup_root=None,
    skip_unlocked=True,
    max_workers=4,
    log_path=None,
    progress_callback=None,
):
    results = []
    total = len(file_list)
    if total == 0:
        return results

    if overwrite and backup_root:
        os.makedirs(backup_root, exist_ok=True)

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futures = {
            ex.submit(
                process_single_pdf,
                file_path,
                common_password,
                per_file_map,
                new_password,
                output_folder,
                overwrite,
                backup_root,
                skip_unlocked,
            ): file_path for file_path in file_list
        }

        completed = 0
        for fut in as_completed(futures):
            file_path = futures[fut]
            try:
                fname, ok, msg = fut.result()
            except Exception as e:
                fname = os.path.basename(file_path)
                ok = False
                msg = f"Unhandled error: {e}"
            completed += 1
            results.append((fname, ok, msg))
            if progress_callback:
                try:
                    progress_callback(completed, total, fname, (ok, msg))
                except Exception:
                    pass

    if log_path:
        try:
            with open(log_path, "w", newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["timestamp", "filename", "success", "message"])
                for r in results:
                    writer.writerow([datetime.now().isoformat(), r[0], "OK" if r[1] else "FAIL", r[2]])
        except Exception as e:
            print(f"Warning: Could not write log file: {e}")

    return results

# -------------------------
# Merge core
# -------------------------
def merge_files_list(file_list, output_path):
    try:
        writer = PdfWriter()
        for file_path in file_list:
            with open(file_path, "rb") as f:
                reader = PdfReader(f)
                for p in reader.pages:
                    writer.add_page(p)
        with open(output_path, "wb") as out_f:
            writer.write(out_f)
        return True, f"Merged {len(file_list)} files into {output_path}"
    except Exception as e:
        return False, f"Error during merge: {e}"

# -------------------------
# Word -> PDF utilities (docx2pdf)
# -------------------------
def convert_docx_file_into_dir(input_path, output_dir):
    """
    Convert a single .docx into output_dir using docx2pdf.
    Returns (ok:bool, message:str, generated_pdf_path or '')
    """
    if not docx2pdf_available:
        return False, "docx2pdf not available", ""
    try:
        os.makedirs(output_dir, exist_ok=True)
        # docx2pdf.convert(input, output) where output is directory or file
        # Passing output_dir will produce basename.pdf inside output_dir
        docx2pdf_convert(input_path, output_dir)
        generated = os.path.join(output_dir, os.path.splitext(os.path.basename(input_path))[0] + ".pdf")
        if os.path.exists(generated):
            return True, "Converted", generated
        else:
            return False, "docx2pdf produced no file", ""
    except Exception as e:
        return False, f"Conversion error: {e}", ""

def convert_docx_folder_into_dir(input_folder, output_folder):
    """
    Convert all docx in input_folder (and subfolders) into output_folder (docx2pdf supports folder->folder).
    Returns (ok, message)
    """
    if not docx2pdf_available:
        return False, "docx2pdf not available"
    try:
        os.makedirs(output_folder, exist_ok=True)
        docx2pdf_convert(input_folder, output_folder)
        return True, "Converted folder"
    except Exception as e:
        return False, f"Folder conversion error: {e}"

# -------------------------
# Main App (Tabbed GUI)
# -------------------------
class TabbedPDFTools(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Tools — Password, Merge, Word→PDF")
        self.geometry("1100x740")
        self.minsize(900, 600)
        self._build_ui()

    def _build_ui(self):
        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True, padx=8, pady=8)

        # Password tool frame
        pwd_frame = ttk.Frame(nb)
        nb.add(pwd_frame, text="Password Tool")
        self._build_password_tab(pwd_frame)

        # Merge tool frame
        merge_frame = ttk.Frame(nb)
        nb.add(merge_frame, text="Merge PDFs")
        self._build_merge_tab(merge_frame)

        # Word -> PDF frame
        word_frame = ttk.Frame(nb)
        nb.add(word_frame, text="Word → PDF")
        self._build_word_tab(word_frame)

    # -------------------------
    # Password tab (unchanged)
    # -------------------------
    def _build_password_tab(self, parent):
        left = ttk.Frame(parent)
        left.pack(side='left', fill='both', expand=True, padx=(4,8), pady=6)

        ttk.Label(left, text="Files to process:").pack(anchor='w')
        self.pwd_listbox = tk.Listbox(left, width=64, height=24, selectmode=tk.EXTENDED)
        self.pwd_listbox.pack(padx=4, pady=6, fill='both', expand=True)

        dnd_text = "Drag & drop PDF files here" if tkdnd_available else "Drag & drop not available — use 'Add files' or 'Add folder'"
        ttk.Label(left, text=dnd_text, foreground='gray').pack(anchor='w', padx=6)

        btn_fr = ttk.Frame(left)
        btn_fr.pack(anchor='w', pady=(4,0))
        ttk.Button(btn_fr, text="Add files", command=self.pwd_add_files).pack(side='left', padx=4)
        ttk.Button(btn_fr, text="Add folder", command=self.pwd_add_folder).pack(side='left', padx=4)
        ttk.Button(btn_fr, text="Remove selected", command=self.pwd_remove_selected).pack(side='left', padx=4)
        ttk.Button(btn_fr, text="Clear", command=self.pwd_clear_list).pack(side='left', padx=4)

        right = ttk.Frame(parent)
        right.pack(side='right', fill='y', padx=(8,4), pady=6)

        ttk.Label(right, text="Common password:").pack(anchor='w')
        self.common_entry = ttk.Entry(right, width=36, show='*')
        self.common_entry.pack(anchor='w', pady=2)

        ttk.Label(right, text="New password (optional):").pack(anchor='w', pady=(6,0))
        self.new_entry = ttk.Entry(right, width=36, show='*')
        self.new_entry.pack(anchor='w', pady=2)

        ttk.Label(right, text="Per-file CSV mapping:").pack(anchor='w', pady=(6,0))
        map_frame = ttk.Frame(right)
        map_frame.pack(anchor='w', pady=2)
        self.map_path_var = tk.StringVar()
        ttk.Entry(map_frame, width=28, textvariable=self.map_path_var).pack(side='left', padx=(0,6))
        ttk.Button(map_frame, text="Load CSV", command=self.pwd_load_map_csv).pack(side='left')

        self.overwrite_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(right, text="Overwrite originals", variable=self.overwrite_var).pack(anchor='w', pady=(6,0))

        ttk.Label(right, text="Output folder (if not overwriting):").pack(anchor='w', pady=(6,0))
        self.output_var = tk.StringVar()
        out_frame = ttk.Frame(right)
        out_frame.pack(anchor='w', pady=2)
        ttk.Entry(out_frame, width=28, textvariable=self.output_var).pack(side='left')
        ttk.Button(out_frame, text="Browse", command=self.pwd_browse_output).pack(side='left', padx=6)

        ttk.Label(right, text="Backup folder (for overwrites):").pack(anchor='w', pady=(6,0))
        self.backup_var = tk.StringVar(value=os.path.join(os.getcwd(), "backups"))
        backup_frame = ttk.Frame(right)
        backup_frame.pack(anchor='w', pady=2)
        ttk.Entry(backup_frame, width=28, textvariable=self.backup_var).pack(side='left')
        ttk.Button(backup_frame, text="Browse", command=self.pwd_browse_backup).pack(side='left', padx=6)

        ttk.Label(right, text="Max worker threads:").pack(anchor='w', pady=(6,0))
        self.workers_spin = ttk.Spinbox(right, from_=1, to=32, width=6)
        self.workers_spin.set(4)
        self.workers_spin.pack(anchor='w', pady=2)

        self.skip_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(right, text="Skip already-unlocked PDFs", variable=self.skip_var).pack(anchor='w', pady=(6,0))

        ttk.Label(right, text="Log file path:").pack(anchor='w', pady=(6,0))
        self.log_var = tk.StringVar(value=os.path.join(os.getcwd(), f"pdf_tool_log_{timestamp_str()}.csv"))
        log_frame = ttk.Frame(right)
        log_frame.pack(anchor='w', pady=2)
        ttk.Entry(log_frame, width=28, textvariable=self.log_var).pack(side='left')
        ttk.Button(log_frame, text="Browse", command=self.pwd_browse_log).pack(side='left', padx=6)

        # Run area bottom
        run_frame = ttk.Frame(parent)
        run_frame.pack(fill='x', padx=12, pady=(6,10))
        self.pwd_run_btn = ttk.Button(run_frame, text="Run", command=self.pwd_start_run)
        self.pwd_run_btn.pack(side='left', padx=(0,8))
        self.pwd_cancel_btn = ttk.Button(run_frame, text="Cancel", command=self.pwd_cancel_run, state=tk.DISABLED)
        self.pwd_cancel_btn.pack(side='left')

        self.pwd_progress_var = tk.DoubleVar(value=0.0)
        self.pwd_progress = ttk.Progressbar(run_frame, variable=self.pwd_progress_var, maximum=100.0, length=520)
        self.pwd_progress.pack(side='left', padx=(8,8))
        self.pwd_status_label = ttk.Label(run_frame, text="Idle")
        self.pwd_status_label.pack(side='left', padx=(6,0))

        if tkdnd_available:
            try:
                self.pwd_listbox.drop_target_register(DND_FILES)
                self.pwd_listbox.dnd_bind('<<Drop>>', self._pwd_on_drop)
            except Exception:
                pass

        self.pwd_filepaths = []

    # Password tab callbacks
    def pwd_add_files(self):
        paths = filedialog.askopenfilenames(title="Select PDF files", filetypes=[("PDF files","*.pdf")])
        if paths:
            for p in paths:
                if p not in self.pwd_filepaths:
                    self.pwd_filepaths.append(p)
                    self.pwd_listbox.insert(tk.END, os.path.basename(p))

    def pwd_add_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing PDFs")
        if folder:
            for root, _, files in os.walk(folder):
                for f in files:
                    if f.lower().endswith('.pdf'):
                        p = os.path.join(root, f)
                        if p not in self.pwd_filepaths:
                            self.pwd_filepaths.append(p)
                            self.pwd_listbox.insert(tk.END, os.path.basename(p))

    def pwd_remove_selected(self):
        sel = list(self.pwd_listbox.curselection())
        for i in reversed(sel):
            self.pwd_listbox.delete(i)
            self.pwd_filepaths.pop(i)

    def pwd_clear_list(self):
        self.pwd_listbox.delete(0, tk.END)
        self.pwd_filepaths = []

    def _pwd_on_drop(self, event):
        data = event.data
        files = self._splitlist_safe(data)
        for f in files:
            f = f.strip('{}')
            if f.lower().endswith('.pdf'):
                if f not in self.pwd_filepaths:
                    self.pwd_filepaths.append(f)
                    self.pwd_listbox.insert(tk.END, os.path.basename(f))

    def pwd_load_map_csv(self):
        path = filedialog.askopenfilename(title="Load per-file CSV", filetypes=[("CSV Files","*.csv"), ("All files","*.*")])
        if path:
            try:
                self.pwd_per_map = load_password_csv(path)
                self.map_path_var.set(path)
                messagebox.showinfo("CSV loaded", f"Loaded mapping for {len(self.pwd_per_map)} files.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load CSV: {e}")

    def pwd_browse_output(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_var.set(path)

    def pwd_browse_backup(self):
        path = filedialog.askdirectory(title="Select backup folder")
        if path:
            self.backup_var.set(path)

    def pwd_browse_log(self):
        path = filedialog.asksaveasfilename(title="Select log file", defaultextension=".csv", filetypes=[("CSV files","*.csv")])
        if path:
            self.log_var.set(path)

    def pwd_start_run(self):
        if not self.pwd_filepaths:
            messagebox.showwarning("No files", "Add PDF files or folders first.")
            return

        common_pw = self.common_entry.get().strip() or None
        new_pw = self.new_entry.get().strip() or None
        overwrite = bool(self.overwrite_var.get())
        output_folder = self.output_var.get().strip() or None
        backup_root = None
        if overwrite:
            backup_root = os.path.join(self.backup_var.get(), f"backup_{timestamp_str()}")
        skip_unlocked = bool(self.skip_var.get())
        try:
            workers = int(self.workers_spin.get())
        except Exception:
            workers = 4
        log_path = self.log_var.get().strip() or os.path.join(os.getcwd(), f"pdf_tool_log_{timestamp_str()}.csv")

        per_map = getattr(self, "pwd_per_map", None)

        if not common_pw and not per_map:
            if not messagebox.askyesno("No password provided", "You didn't provide a common password or per-file map. Processing will rewrite files as-is (if unlocked). Continue?"):
                return

        # disable UI
        self.pwd_run_btn.config(state=tk.DISABLED)
        self.pwd_cancel_btn.config(state=tk.NORMAL)
        self.pwd_running = True
        self.pwd_status_label.config(text="Running...")
        self.pwd_progress_var.set(0.0)

        def background_job():
            def progress_cb(completed, total, filename, status):
                perc = (completed / total) * 100.0
                self._pwd_queue.put(("progress", perc, filename, status))

            try:
                results = run_batch(
                    self.pwd_filepaths.copy(),
                    common_password=common_pw,
                    per_file_map=per_map,
                    new_password=new_pw,
                    output_folder=output_folder,
                    overwrite=overwrite,
                    backup_root=backup_root,
                    skip_unlocked=skip_unlocked,
                    max_workers=workers,
                    log_path=log_path,
                    progress_callback=progress_cb,
                )
                self._pwd_queue.put(("done", results))
            except Exception as e:
                tb = traceback.format_exc()
                self._pwd_queue.put(("error", f"{e}\n{tb}"))

        self._pwd_queue = Queue()
        threading.Thread(target=background_job, daemon=True).start()
        self.after(150, self._pwd_process_queue)

    def pwd_cancel_run(self):
        if messagebox.askyesno("Cancel", "Cancel requested. Running threads may finish current file, but UI will stop tracking progress. Continue?"):
            self.pwd_running = False
            self.pwd_run_btn.config(state=tk.NORMAL)
            self.pwd_cancel_btn.config(state=tk.DISABLED)
            self.pwd_status_label.config(text="Cancelled")

    def _pwd_process_queue(self):
        try:
            while True:
                item = self._pwd_queue.get_nowait()
                if item[0] == "progress":
                    _, perc, filename, status = item
                    self.pwd_progress_var.set(perc)
                    ok, msg = status
                    self.pwd_status_label.config(text=f"{filename} → {'OK' if ok else 'FAIL'}")
                elif item[0] == "done":
                    _, results = item
                    ok_count = sum(1 for r in results if r[1])
                    self.pwd_status_label.config(text=f"Done: {ok_count}/{len(results)} succeeded")
                    self.pwd_progress_var.set(100.0)
                    messagebox.showinfo("Completed", f"Completed: {ok_count}/{len(results)} succeeded.\nLog: {self.log_var.get()}")
                    self.pwd_run_btn.config(state=tk.NORMAL)
                    self.pwd_cancel_btn.config(state=tk.DISABLED)
                    self.pwd_running = False
                elif item[0] == "error":
                    _, msg = item
                    messagebox.showerror("Error", msg)
                    self.pwd_run_btn.config(state=tk.NORMAL)
                    self.pwd_cancel_btn.config(state=tk.DISABLED)
                    self.pwd_running = False
        except Empty:
            pass

        if getattr(self, "pwd_running", False):
            self.after(150, self._pwd_process_queue)
        else:
            self.pwd_run_btn.config(state=tk.NORMAL)
            self.pwd_cancel_btn.config(state=tk.DISABLED)

    # -------------------------
    # Merge tab (Add Folder included)
    # -------------------------
    def _build_merge_tab(self, parent):
        top = ttk.Frame(parent)
        top.pack(fill='both', expand=True, padx=6, pady=6)

        ttk.Label(top, text="Files to merge:").pack(anchor='w')
        self.merge_listbox = tk.Listbox(top, width=80, height=20, selectmode=tk.EXTENDED)
        self.merge_listbox.pack(padx=4, pady=6, fill='both', expand=True)

        btn_fr = ttk.Frame(top)
        btn_fr.pack(anchor='w', pady=(2,0))
        ttk.Button(btn_fr, text="Add files", command=self.merge_add_files).pack(side='left', padx=4)
        ttk.Button(btn_fr, text="Add folder", command=self.merge_add_folder).pack(side='left', padx=4)
        ttk.Button(btn_fr, text="Remove selected", command=self.merge_remove_selected).pack(side='left', padx=4)
        ttk.Button(btn_fr, text="Clear", command=self.merge_clear_list).pack(side='left', padx=4)
        ttk.Button(btn_fr, text="Move Up", command=self.merge_move_up).pack(side='left', padx=4)
        ttk.Button(btn_fr, text="Move Down", command=self.merge_move_down).pack(side='left', padx=4)

        out_frame = ttk.Frame(top)
        out_frame.pack(fill='x', pady=(8,0))
        ttk.Label(out_frame, text="Output file:").pack(anchor='w')
        self.merge_output_var = tk.StringVar(value=os.path.join(os.getcwd(), f"merged_{timestamp_str()}.pdf"))
        out_sub = ttk.Frame(out_frame)
        out_sub.pack(anchor='w', pady=2)
        ttk.Entry(out_sub, width=56, textvariable=self.merge_output_var).pack(side='left')
        ttk.Button(out_sub, text="Browse", command=self.merge_browse_output).pack(side='left', padx=6)

        ttk.Button(top, text="Merge", command=self.merge_run).pack(anchor='w', pady=(10,0))

        if tkdnd_available:
            try:
                self.merge_listbox.drop_target_register(DND_FILES)
                self.merge_listbox.dnd_bind('<<Drop>>', self._merge_on_drop)
            except Exception:
                pass

        self.merge_filepaths = []

    def _merge_on_drop(self, event):
        data = event.data
        files = self._splitlist_safe(data)
        for f in files:
            f = f.strip('{}')
            if f.lower().endswith('.pdf'):
                if f not in self.merge_filepaths:
                    self.merge_filepaths.append(f)
                    self.merge_listbox.insert(tk.END, os.path.basename(f))

    def merge_add_files(self):
        paths = filedialog.askopenfilenames(title="Select PDF files", filetypes=[("PDF files","*.pdf")])
        if paths:
            for p in paths:
                if p not in self.merge_filepaths:
                    self.merge_filepaths.append(p)
                    self.merge_listbox.insert(tk.END, os.path.basename(p))

    def merge_add_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing PDFs to add")
        if folder:
            added = 0
            for root, _, files in os.walk(folder):
                for f in files:
                    if f.lower().endswith('.pdf'):
                        p = os.path.join(root, f)
                        if p not in self.merge_filepaths:
                            self.merge_filepaths.append(p)
                            self.merge_listbox.insert(tk.END, os.path.basename(p))
                            added += 1
            messagebox.showinfo("Folder scanned", f"Added {added} PDF files from folder.")

    def merge_remove_selected(self):
        sel = list(self.merge_listbox.curselection())
        for i in reversed(sel):
            self.merge_listbox.delete(i)
            self.merge_filepaths.pop(i)

    def merge_clear_list(self):
        self.merge_listbox.delete(0, tk.END)
        self.merge_filepaths = []

    def merge_move_up(self):
        sel = list(self.merge_listbox.curselection())
        for i in sel:
            if i == 0:
                continue
            self.merge_filepaths[i], self.merge_filepaths[i-1] = self.merge_filepaths[i-1], self.merge_filepaths[i]
            name1 = os.path.basename(self.merge_filepaths[i-1])
            name2 = os.path.basename(self.merge_filepaths[i])
            self.merge_listbox.delete(i-1, i)
            self.merge_listbox.insert(i-1, name1)
            self.merge_listbox.insert(i, name2)
            self.merge_listbox.select_clear(0, tk.END)
            self.merge_listbox.select_set(i-1)

    def merge_move_down(self):
        sel = list(self.merge_listbox.curselection())
        for i in reversed(sel):
            if i == self.merge_listbox.size() - 1:
                continue
            self.merge_filepaths[i], self.merge_filepaths[i+1] = self.merge_filepaths[i+1], self.merge_filepaths[i]
            name1 = os.path.basename(self.merge_filepaths[i])
            name2 = os.path.basename(self.merge_filepaths[i+1])
            self.merge_listbox.delete(i, i+1)
            self.merge_listbox.insert(i, name1)
            self.merge_listbox.insert(i+1, name2)
            self.merge_listbox.select_clear(0, tk.END)
            self.merge_listbox.select_set(i+1)

    def merge_browse_output(self):
        path = filedialog.asksaveasfilename(title="Save merged PDF as", defaultextension=".pdf", filetypes=[("PDF files","*.pdf")])
        if path:
            self.merge_output_var.set(path)

    def merge_run(self):
        if not self.merge_filepaths:
            messagebox.showwarning("No files", "Add PDF files to merge first.")
            return
        out_path = self.merge_output_var.get().strip()
        if not out_path:
            messagebox.showwarning("No output", "Select an output file path.")
            return
        ok, msg = merge_files_list(self.merge_filepaths, out_path)
        if ok:
            messagebox.showinfo("Merged", msg)
        else:
            messagebox.showerror("Error", msg)

    # -------------------------
    # Word -> PDF tab (enhanced)
    # -------------------------
    def _build_word_tab(self, parent):
        top = ttk.Frame(parent)
        top.pack(fill='both', expand=True, padx=6, pady=6)

        header_fr = ttk.Frame(top)
        header_fr.pack(fill='x')
        ttk.Label(header_fr, text="Word (.docx) files to convert:", font=("TkDefaultFont", 10, "bold")).pack(anchor='w')

        self.word_listbox = tk.Listbox(top, width=80, height=8, selectmode=tk.EXTENDED)
        self.word_listbox.pack(padx=4, pady=6, fill='x', expand=False)

        btn_fr = ttk.Frame(top)
        btn_fr.pack(anchor='w', pady=(2,0))
        ttk.Button(btn_fr, text="Add .docx files", command=self.word_add_files).pack(side='left', padx=4)
        ttk.Button(btn_fr, text="Add folder", command=self.word_add_folder).pack(side='left', padx=4)
        ttk.Button(btn_fr, text="Remove selected", command=self.word_remove_selected).pack(side='left', padx=4)
        ttk.Button(btn_fr, text="Clear", command=self.word_clear_list).pack(side='left', padx=4)

        # Preserve subfolders option
        opt_fr = ttk.Frame(top)
        opt_fr.pack(anchor='w', pady=(6,0))
        self.preserve_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_fr, text="Preserve subfolders when adding folders", variable=self.preserve_var).pack(side='left', padx=(0,8))
        ttk.Label(opt_fr, text="(If unchecked, all PDFs go into the chosen output folder.)", foreground='gray').pack(side='left')

        out_frame = ttk.Frame(top)
        out_frame.pack(fill='x', pady=(8,0))
        ttk.Label(out_frame, text="Output folder (PDFs will be placed here):").pack(anchor='w')
        self.word_output_var = tk.StringVar(value=os.path.join(os.getcwd(), "word_to_pdf_output"))
        out_sub = ttk.Frame(out_frame)
        out_sub.pack(anchor='w', pady=2)
        ttk.Entry(out_sub, width=56, textvariable=self.word_output_var).pack(side='left')
        ttk.Button(out_sub, text="Browse", command=self.word_browse_output).pack(side='left', padx=6)
        ttk.Button(out_sub, text="Open Output Folder", command=self.word_open_output_folder).pack(side='left', padx=6)

        run_fr = ttk.Frame(top)
        run_fr.pack(fill='x', pady=(10,0))
        self.word_convert_btn = ttk.Button(run_fr, text="Convert Selected", command=self.word_convert_selected)
        self.word_convert_btn.pack(side='left', padx=(0,8))
        self.word_convert_all_btn = ttk.Button(run_fr, text="Convert All (in list)", command=self.word_convert_all)
        self.word_convert_all_btn.pack(side='left', padx=(0,8))

        self.word_progress_var = tk.DoubleVar(value=0.0)
        self.word_progress = ttk.Progressbar(run_fr, variable=self.word_progress_var, maximum=100.0, length=420)
        self.word_progress.pack(side='left', padx=(8,8))
        self.word_status_label = ttk.Label(run_fr, text="Idle")
        self.word_status_label.pack(side='left', padx=(6,0))

        # Results Treeview
        result_fr = ttk.Frame(top)
        result_fr.pack(fill='both', expand=True, pady=(10,0))
        cols = ("file", "status", "message", "output")
        self.word_tree = ttk.Treeview(result_fr, columns=cols, show='headings', height=10)
        self.word_tree.heading('file', text='File')
        self.word_tree.heading('status', text='Status')
        self.word_tree.heading('message', text='Message')
        self.word_tree.heading('output', text='Output Path')
        self.word_tree.column('file', width=220)
        self.word_tree.column('status', width=80)
        self.word_tree.column('message', width=200)
        self.word_tree.column('output', width=380)
        self.word_tree.pack(fill='both', expand=True, side='left')

        scrollbar = ttk.Scrollbar(result_fr, orient='vertical', command=self.word_tree.yview)
        self.word_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side='right', fill='y')

        if tkdnd_available:
            try:
                self.word_listbox.drop_target_register(DND_FILES)
                self.word_listbox.dnd_bind('<<Drop>>', self._word_on_drop)
            except Exception:
                pass

        # mapping of file -> origin folder (if added from folder), used for preserving subfolders
        self.word_filepaths = []
        self.word_origin_map = {}  # file_path -> origin_folder (or None)

        # Disable convert buttons if docx2pdf unavailable
        if not docx2pdf_available:
            self.word_convert_btn.config(state=tk.DISABLED)
            self.word_convert_all_btn.config(state=tk.DISABLED)
            ttk.Label(top, text="docx2pdf or Microsoft Word not available. Install docx2pdf and ensure MS Word is installed on Windows.", foreground='red').pack(anchor='w', pady=(6,0))

    # Word tab callbacks
    def _word_on_drop(self, event):
        data = event.data
        files = self._splitlist_safe(data)
        for f in files:
            f = f.strip('{}')
            if f.lower().endswith('.docx'):
                if f not in self.word_filepaths:
                    self.word_filepaths.append(f)
                    self.word_origin_map[f] = None
                    self.word_listbox.insert(tk.END, os.path.basename(f))

    def word_add_files(self):
        paths = filedialog.askopenfilenames(title="Select .docx files", filetypes=[("Word files","*.docx")])
        if paths:
            for p in paths:
                if p not in self.word_filepaths:
                    self.word_filepaths.append(p)
                    self.word_origin_map[p] = None
                    self.word_listbox.insert(tk.END, os.path.basename(p))

    def word_add_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing .docx files")
        if folder:
            added = 0
            for root, _, files in os.walk(folder):
                for f in files:
                    if f.lower().endswith('.docx'):
                        p = os.path.join(root, f)
                        if p not in self.word_filepaths:
                            self.word_filepaths.append(p)
                            # store origin folder to compute relative path if preserve is enabled
                            self.word_origin_map[p] = folder
                            self.word_listbox.insert(tk.END, os.path.basename(p))
                            added += 1
            messagebox.showinfo("Folder scanned", f"Added {added} .docx files from folder.")

    def word_remove_selected(self):
        sel = list(self.word_listbox.curselection())
        for i in reversed(sel):
            name = self.word_listbox.get(i)
            # remove by index
            # need to find the full path corresponding to this index
            # assume insertion order preserved
            fullpath = self.word_filepaths[i]
            self.word_listbox.delete(i)
            self.word_filepaths.pop(i)
            self.word_origin_map.pop(fullpath, None)
            # also remove from tree if present
            for iid in self.word_tree.get_children():
                vals = self.word_tree.item(iid, 'values')
                if vals and vals[0] == os.path.basename(fullpath):
                    self.word_tree.delete(iid)

    def word_clear_list(self):
        self.word_listbox.delete(0, tk.END)
        self.word_filepaths = []
        self.word_origin_map = {}
        for iid in self.word_tree.get_children():
            self.word_tree.delete(iid)

    def word_browse_output(self):
        path = filedialog.askdirectory(title="Select output folder for PDFs")
        if path:
            self.word_output_var.set(path)

    def word_open_output_folder(self):
        out = self.word_output_var.get().strip() or os.getcwd()
        if not open_folder_in_explorer(out):
            messagebox.showwarning("Open folder", f"Could not open folder: {out}")

    def word_convert_selected(self):
        sel = list(self.word_listbox.curselection())
        if not sel:
            messagebox.showwarning("No selection", "Select files to convert.")
            return
        to_convert = [self.word_filepaths[i] for i in sel]
        self._start_word_conversion(to_convert)

    def word_convert_all(self):
        if not self.word_filepaths:
            messagebox.showwarning("No files", "Add Word files to convert first.")
            return
        self._start_word_conversion(self.word_filepaths.copy())

    def _start_word_conversion(self, file_list):
        if not docx2pdf_available:
            messagebox.showerror("Unavailable", "docx2pdf not installed or Microsoft Word not available.")
            return
        out_dir = self.word_output_var.get().strip() or os.getcwd()
        preserve = bool(self.preserve_var.get())

        # clear results table for these files (append new rows)
        for fp in file_list:
            # insert a row with status "Queued"
            self.word_tree.insert("", "end", values=(os.path.basename(fp), "Queued", "", ""))

        # disable UI
        self.word_convert_btn.config(state=tk.DISABLED)
        self.word_convert_all_btn.config(state=tk.DISABLED)
        self.word_status_label.config(text="Converting...")
        self.word_progress_var.set(0.0)

        q = Queue()

        def background_job():
            total = len(file_list)
            completed = 0
            results = []
            for fp in file_list:
                try:
                    origin = self.word_origin_map.get(fp, None)
                    if preserve and origin:
                        # compute relative path to origin, then use that relative dir under out_dir
                        rel = os.path.relpath(fp, origin)
                        rel_dir = os.path.dirname(rel)  # may be ''
                        target_dir = os.path.join(out_dir, rel_dir) if rel_dir else out_dir
                        ok, msg, generated = convert_docx_file_into_dir(fp, target_dir)
                        out_path = generated if ok else ""
                    else:
                        # convert into out_dir directly
                        ok, msg, generated = convert_docx_file_into_dir(fp, out_dir)
                        out_path = generated if ok else ""
                except Exception as e:
                    ok = False
                    msg = f"Error: {e}"
                    out_path = ""
                completed += 1
                pct = (completed / total) * 100.0
                q.put(("progress", pct, fp, ok, msg, out_path))
                results.append((fp, ok, msg, out_path))
            q.put(("done", results))

        threading.Thread(target=background_job, daemon=True).start()

        def poll():
            try:
                while True:
                    item = q.get_nowait()
                    if item[0] == "progress":
                        _, pct, fp, ok, msg, outp = item
                        self.word_progress_var.set(pct)
                        self.word_status_label.config(text=f"{os.path.basename(fp)} → {'OK' if ok else 'FAIL'}")
                        # update tree: find first row with filename basename(fp) and status "Queued" OR first match without output
                        updated = False
                        for iid in self.word_tree.get_children():
                            vals = self.word_tree.item(iid, 'values')
                            if vals and vals[0] == os.path.basename(fp) and (vals[1] in ("Queued", "In Progress", "") or vals[3] == ""):
                                status_text = "OK" if ok else "FAIL"
                                self.word_tree.item(iid, values=(vals[0], status_text, msg, outp))
                                updated = True
                                break
                        if not updated:
                            # append
                            status_text = "OK" if ok else "FAIL"
                            self.word_tree.insert("", "end", values=(os.path.basename(fp), status_text, msg, outp))
                    elif item[0] == "done":
                        _, results = item
                        ok_count = sum(1 for r in results if r[1])
                        self.word_progress_var.set(100.0)
                        self.word_status_label.config(text=f"Done: {ok_count}/{len(results)} succeeded")
                        messagebox.showinfo("Word → PDF", f"Converted {ok_count}/{len(results)} files.\nOutput folder: {out_dir}")
                        self.word_convert_btn.config(state=tk.NORMAL)
                        self.word_convert_all_btn.config(state=tk.NORMAL)
                        # enable Open Output Folder (already present)
            except Empty:
                pass
            # continue polling until progress == 100
            if self.word_progress_var.get() < 100.0:
                self.after(150, poll)
            else:
                self.word_convert_btn.config(state=tk.NORMAL)
                self.word_convert_all_btn.config(state=tk.NORMAL)
                self.word_status_label.config(text="Idle")

        self.after(150, poll)

    # -------------------------
    # Shared helpers
    # -------------------------
    def _splitlist_safe(self, data):
        try:
            return self.tk.splitlist(data)
        except Exception:
            return data.split()

# -------------------------
# README & PyInstaller spec (printed here for convenience)
# -------------------------
README_TEXT = r"""
pdf_tools_tabbed_word_improved.py — README
-----------------------------------------

This single-file GUI app contains three tools:
 - Password Tool (PDF decrypt / re-encrypt / batch)
 - Merge PDFs (reorder + merge)
 - Word -> PDF (docx -> pdf using docx2pdf, requires Microsoft Word on Windows)

Requirements:
 - Python 3.8+
 - pip install PyPDF2 pandas docx2pdf tkinterdnd2
   - pandas is optional (CSV convenience)
   - tkinterdnd2 is optional (drag & drop in the GUI)
 - For Word->PDF: Microsoft Word must be installed (Windows). docx2pdf uses Word automation.

Quick start:
1. Install dependencies:
   pip install PyPDF2 pandas docx2pdf tkinterdnd2

2. Run:
   python pdf_tools_tabbed_word_improved.py

Word -> PDF notes:
 - Use "Add folder" to add all .docx files from a folder.
 - Check "Preserve subfolders when adding folders" to keep relative folder structure inside the chosen output folder.
 - You can "Convert Selected" or "Convert All".
 - The results table shows status and output path.
 - Click "Open Output Folder" to reveal results.

Building a Windows EXE with PyInstaller:
----------------------------------------
Below is a minimal PyInstaller recipe. Save as e.g. build.spec and run `pyinstaller build.spec`.

Note: docx2pdf calls Word via COM; the resulting EXE must be run on a machine with MS Word installed.

.spec example (see below in the code comments):
 - single-file exe
 - includes data files if needed
 - tweak the icon and hiddenimports as required

Security note:
 - Be careful when overwriting originals; use the backup option in the Password tab.

"""

PYINSTALLER_SPEC = r"""
# PyInstaller spec example for pdf_tools_tabbed_word_improved.py
# Save this text to pdf_tools.spec and run:
#   pyinstaller --onefile pdf_tools.spec

# This spec is a starting point. You may need to add hiddenimports or adjust paths.

# Replace 'your_icon.ico' with your icon path if desired.

block_cipher = None

a = Analysis(
    ['pdf_tools_tabbed_word_improved.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['docx2pdf'],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='pdf_tools',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='your_icon.ico'  # optional
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='pdf_tools_dist'
)
"""

# -------------------------
# Run app
# -------------------------
if __name__ == "__main__":
    app = TabbedPDFTools()
    app.mainloop()
