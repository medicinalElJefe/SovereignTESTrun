#!/usr/bin/env python3
"""
Sovereign Doc – GUI
"""

import os
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

from sovereign_doc_core import convert_any, SovereignDocError


class SovereignDocApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sovereign Doc – Round Trip")
        self.geometry("700x320")
        self.resizable(False, False)
        self._build_widgets()

    def get_version(self) -> str:
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            version_path = os.path.join(base_dir, "VERSION")
            with open(version_path, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception:
            return "Unknown"

    def _build_widgets(self):
        menubar = tk.Menu(self)
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About Sovereign Doc", command=self.on_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        self.config(menu=menubar)

        file_frame = tk.LabelFrame(self, text="Single File", padx=8, pady=6)
        file_frame.pack(fill="x", padx=12, pady=(12, 6))

        tk.Label(file_frame, text="Input file (.docx, .txt, .md):").pack(anchor="w")

        path_frame = tk.Frame(file_frame)
        path_frame.pack(fill="x", pady=(2, 0))

        self.path_var = tk.StringVar()
        entry = tk.Entry(path_frame, textvariable=self.path_var)
        entry.pack(side="left", fill="x", expand=True)

        browse_btn = tk.Button(path_frame, text="Browse…", command=self.on_browse_file)
        browse_btn.pack(side="left", padx=(6, 0))

        fmt_frame = tk.Frame(file_frame)
        fmt_frame.pack(fill="x", pady=(8, 0))

        tk.Label(fmt_frame, text="Output format:").pack(anchor="w")

        self.format_var = tk.StringVar(value="txt")
        for label, val in [
            ("Plain Text (.txt)", "txt"),
            ("Markdown (.md)", "md"),
            ("HTML (.html)", "html"),
            ("Word (.docx)", "docx"),
        ]:
            rb = tk.Radiobutton(fmt_frame, text=label, variable=self.format_var, value=val)
            rb.pack(anchor="w")

        self.log_var = tk.BooleanVar(value=True)
        cb_log = tk.Checkbutton(file_frame, text="Log conversions to sovereign_doc_log.csv", variable=self.log_var)
        cb_log.pack(anchor="w", pady=(4, 0))

        single_btn = tk.Button(file_frame, text="Convert This File", command=self.on_convert_single, width=20)
        single_btn.pack(anchor="w", pady=(8, 0))

        folder_frame = tk.LabelFrame(self, text="Folder Batch (.docx only)", padx=8, pady=6)
        folder_frame.pack(fill="x", padx=12, pady=(6, 6))

        tk.Label(folder_frame, text="Folder with .docx files:").pack(anchor="w")

        folder_path_frame = tk.Frame(folder_frame)
        folder_path_frame.pack(fill="x", pady=(2, 0))

        self.folder_var = tk.StringVar()
        folder_entry = tk.Entry(folder_path_frame, textvariable=self.folder_var)
        folder_entry.pack(side="left", fill="x", expand=True)

        browse_folder_btn = tk.Button(folder_path_frame, text="Browse Folder…", command=self.on_browse_folder)
        browse_folder_btn.pack(side="left", padx=(6, 0))

        batch_btn = tk.Button(folder_frame, text="Convert All in Folder", command=self.on_convert_folder, width=22)
        batch_btn.pack(anchor="w", pady=(8, 0))

        self.status_var = tk.StringVar(value="Ready.")
        status_label = tk.Label(self, textvariable=self.status_var, anchor="w", fg="#222222")
        status_label.pack(fill="x", padx=12, pady=(8, 4))

    def on_about(self):
        version = self.get_version()
        messagebox.showinfo(
            "About Sovereign Doc",
            f"Sovereign Doc v{version}\n\n"
            "Local-only document liberation and round-trip tool.\n"
            "No cloud. No telemetry. No subscriptions.\n\n"
            "Format freedom = digital sovereignty.\n"
            "Created by Jeffrey Alan Dewey."
        )

    def on_browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select file (.docx, .txt, .md)",
            filetypes=[
                ("Supported documents", "*.docx *.txt *.md"),
                ("Word documents", "*.docx"),
                ("Text files", "*.txt"),
                ("Markdown files", "*.md"),
                ("All files", "*.*"),
            ]
        )
        if filename:
            self.path_var.set(filename)
            self.status_var.set("Selected file.")

    def on_browse_folder(self):
        foldername = filedialog.askdirectory(title="Select folder containing .docx files")
        if foldername:
            self.folder_var.set(foldername)
            self.status_var.set("Selected folder.")

    def on_convert_single(self):
        input_path_str = self.path_var.get().strip()
        if not input_path_str:
            messagebox.showwarning("Sovereign Doc", "Please select an input file first.")
            return
        input_path = Path(input_path_str).expanduser().resolve()
        if not input_path.is_file():
            messagebox.showerror("Sovereign Doc", f"File not found:\n{input_path}")
            return
        dst_fmt = self.format_var.get()
        logging_enabled = bool(self.log_var.get())
        try:
            out_path = convert_any(input_path, dst_fmt, mode="gui-single", enable_log=logging_enabled)
        except SovereignDocError as e:
            self.status_var.set("Conversion failed.")
            messagebox.showerror("Sovereign Doc – Error", str(e))
            return
        except Exception as e:
            self.status_var.set("Conversion failed (internal error).")
            tb = traceback.format_exc()
            messagebox.showerror("Sovereign Doc – Fatal Error", f"{e}\n\n{tb}")
            return
        self.status_var.set(f"Converted → {out_path.name}")
        messagebox.showinfo(
            "Sovereign Doc",
            f"Conversion complete:\n{input_path.name} → {out_path.name}\n\nSaved at:\n{out_path}"
        )

    def on_convert_folder(self):
        folder_str = self.folder_var.get().strip()
        if not folder_str:
            messagebox.showwarning("Sovereign Doc", "Please select a folder first.")
            return
        folder_path = Path(folder_str).expanduser().resolve()
        if not folder_path.is_dir():
            messagebox.showerror("Sovereign Doc", f"Folder not found:\n{folder_path}")
            return
        dst_fmt = self.format_var.get()
        logging_enabled = bool(self.log_var.get())
        docx_files = sorted(folder_path.glob("*.docx"))
        if not docx_files:
            messagebox.showinfo("Sovereign Doc", f"No .docx files found in:\n{folder_path}")
            return
        count = 0
        errors = 0
        for f in docx_files:
            try:
                convert_any(f, dst_fmt, mode="gui-batch", enable_log=logging_enabled)
            except Exception:
                errors += 1
                continue
            else:
                count += 1
        self.status_var.set(f"Batch complete: {count} converted, {errors} failed.")
        messagebox.showinfo(
            "Sovereign Doc – Batch Complete",
            f"Folder: {folder_path}\nConverted: {count}\nFailed: {errors}"
        )


def main():
    app = SovereignDocApp()
    app.mainloop()


if __name__ == "__main__":
    main()
