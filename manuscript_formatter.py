#!/usr/bin/env python3
"""
Manuscript Formatter
====================
A unified GUI tool that reformats .docx manuscripts into journal-specific
formats. Select a manuscript, pick a target format, and save the output.

Features:
    - Upload .ris bibliography file to auto-format references
    - Optional Zotero field code embedding for citation management
    - Word-recognizable SEQ field captions

Usage:
    python manuscript_formatter.py
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from reader import read_manuscript
from formats import get_formats
from ris_parser import parse_ris


class ManuscriptFormatterApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Manuscript Formatter')
        self.root.resizable(False, False)

        self.manuscript_path = None
        self.ris_path = None
        self.formats = get_formats()

        if not self.formats:
            messagebox.showerror('Error', 'No format plugins found.')
            sys.exit(1)

        self._build_ui()
        self._center_window()

    def _center_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f'+{x}+{y}')

    def _build_ui(self):
        frame = tk.Frame(self.root, padx=20, pady=15)
        frame.pack()

        tk.Label(frame, text='Manuscript Formatter',
                 font=('Segoe UI', 14, 'bold')).pack(pady=(0, 15))

        # ── Manuscript file ──
        file_frame = tk.Frame(frame)
        file_frame.pack(fill='x', pady=(0, 5))
        tk.Label(file_frame, text='Manuscript (.docx):',
                 font=('Segoe UI', 10)).pack(anchor='w')

        path_frame = tk.Frame(file_frame)
        path_frame.pack(fill='x', pady=(2, 0))
        self.path_var = tk.StringVar(value='No file selected')
        tk.Label(path_frame, textvariable=self.path_var,
                 font=('Segoe UI', 9), fg='#666',
                 anchor='w', width=50).pack(side='left', fill='x', expand=True)
        tk.Button(path_frame, text='Open File...',
                  command=self._open_file,
                  font=('Segoe UI', 9)).pack(side='right', padx=(10, 0))

        # ── RIS bibliography file (optional) ──
        ris_frame = tk.Frame(frame)
        ris_frame.pack(fill='x', pady=(5, 5))
        tk.Label(ris_frame, text='Bibliography (.ris) — optional:',
                 font=('Segoe UI', 10)).pack(anchor='w')

        ris_path_frame = tk.Frame(ris_frame)
        ris_path_frame.pack(fill='x', pady=(2, 0))
        self.ris_var = tk.StringVar(value='No file selected')
        tk.Label(ris_path_frame, textvariable=self.ris_var,
                 font=('Segoe UI', 9), fg='#666',
                 anchor='w', width=50).pack(side='left', fill='x', expand=True)

        ris_btn_frame = tk.Frame(ris_path_frame)
        ris_btn_frame.pack(side='right')
        tk.Button(ris_btn_frame, text='Open .ris...',
                  command=self._open_ris,
                  font=('Segoe UI', 9)).pack(side='left', padx=(10, 0))
        tk.Button(ris_btn_frame, text='Clear',
                  command=self._clear_ris,
                  font=('Segoe UI', 9)).pack(side='left', padx=(5, 0))

        # ── Format selection ──
        fmt_frame = tk.LabelFrame(frame, text='Format',
                                  font=('Segoe UI', 10), padx=10, pady=5)
        fmt_frame.pack(fill='x', pady=(5, 5))
        self.format_var = tk.StringVar()
        format_names = sorted(self.formats.keys())
        self.format_var.set(format_names[0])
        for name in format_names:
            tk.Radiobutton(fmt_frame, text=name, variable=self.format_var,
                           value=name, font=('Segoe UI', 10)
                           ).pack(side='left', padx=(0, 15))

        # ── Options ──
        opt_frame = tk.LabelFrame(frame, text='Options',
                                  font=('Segoe UI', 10), padx=10, pady=5)
        opt_frame.pack(fill='x', pady=(5, 10))

        self.zotero_var = tk.BooleanVar(value=False)
        tk.Checkbutton(opt_frame,
                       text='Embed Zotero field codes in references',
                       variable=self.zotero_var,
                       font=('Segoe UI', 9)).pack(anchor='w')

        self.crossref_var = tk.BooleanVar(value=False)
        tk.Checkbutton(opt_frame,
                       text='Look up unmatched references via CrossRef (requires internet)',
                       variable=self.crossref_var,
                       font=('Segoe UI', 9)).pack(anchor='w')

        # ── Format button ──
        tk.Button(frame, text='Format Manuscript',
                  command=self._format_manuscript,
                  font=('Segoe UI', 11, 'bold'),
                  bg='#0078D4', fg='white', padx=20, pady=5
                  ).pack(pady=(5, 10))

        # ── Status ──
        self.status_var = tk.StringVar(value='Ready')
        tk.Label(frame, textvariable=self.status_var,
                 font=('Segoe UI', 9), fg='#888').pack()

    def _open_file(self):
        path = filedialog.askopenfilename(
            title='Select your manuscript (.docx)',
            filetypes=[('Word Documents', '*.docx'), ('All Files', '*.*')],
        )
        if path:
            self.manuscript_path = os.path.abspath(path)
            self.path_var.set(os.path.basename(self.manuscript_path))
            self.status_var.set('Ready')

    def _open_ris(self):
        path = filedialog.askopenfilename(
            title='Select bibliography file (.ris)',
            filetypes=[('RIS Files', '*.ris'), ('All Files', '*.*')],
        )
        if path:
            self.ris_path = os.path.abspath(path)
            self.ris_var.set(os.path.basename(self.ris_path))
            self.status_var.set('Ready')

    def _clear_ris(self):
        self.ris_path = None
        self.ris_var.set('No file selected')

    def _format_manuscript(self):
        if not self.manuscript_path or not os.path.isfile(self.manuscript_path):
            messagebox.showwarning('No File',
                                   'Please select a manuscript file first.')
            return

        format_name = self.format_var.get()
        plugin = self.formats[format_name]

        # Read manuscript
        self.status_var.set('Reading manuscript...')
        self.root.update()
        try:
            items = read_manuscript(self.manuscript_path)
        except Exception as e:
            messagebox.showerror('Read Error',
                                 f'Failed to read manuscript:\n{e}')
            self.status_var.set('Error')
            return

        # Parse RIS if provided
        ris_data = None
        if self.ris_path and os.path.isfile(self.ris_path):
            self.status_var.set('Parsing bibliography...')
            self.root.update()
            try:
                ris_data = parse_ris(self.ris_path)
            except Exception as e:
                messagebox.showwarning('RIS Warning',
                                       f'Could not parse .ris file:\n{e}\n\n'
                                       'Continuing without bibliography data.')
                ris_data = None

        # Save As dialog
        manuscript_name = os.path.splitext(
            os.path.basename(self.manuscript_path))[0]
        default_name = f'{manuscript_name}{plugin.FORMAT_SUFFIX}.docx'
        default_dir = os.path.dirname(self.manuscript_path)

        output_path = filedialog.asksaveasfilename(
            title='Save formatted manuscript as...',
            defaultextension='.docx',
            filetypes=[('Word Documents', '*.docx')],
            initialfile=default_name,
            initialdir=default_dir,
        )
        if not output_path:
            self.status_var.set('Cancelled')
            return

        # Progress callback for GUI updates
        def progress_cb(current, total, message):
            self.status_var.set(message)
            self.root.update()

        # Build
        self.status_var.set(f'Building {format_name} document...')
        self.root.update()
        try:
            plugin.build(items, output_path,
                         ris_data=ris_data,
                         zotero_enabled=self.zotero_var.get(),
                         use_crossref=self.crossref_var.get(),
                         progress_callback=progress_cb)
        except Exception as e:
            messagebox.showerror('Build Error',
                                 f'Failed to build document:\n{e}')
            self.status_var.set('Error')
            return

        self.status_var.set('Done!')
        msg = f'Saved to:\n{output_path}'
        if ris_data:
            msg += f'\n\nMatched references against {len(ris_data)} RIS records.'
        messagebox.showinfo('Success', msg)


def main():
    root = tk.Tk()
    root.attributes('-topmost', True)
    app = ManuscriptFormatterApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
