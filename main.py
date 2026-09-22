"""TestLink -> WE/TelecomEgypt formatter: desktop GUI (plain Tkinter / ttk).

All real work lives in skill_core.py and runs on a background thread so the
window never freezes. Progress and errors come back through a queue that the
main thread drains with .after() (widgets are only touched from the main thread).

Only the Python standard library is used for the GUI: no third-party GUI package,
so the packaged app stays small and simple.
"""
import os
import queue
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import skill_core

APP_TITLE = "TestLink to WE Formatter"
RUN_TEXT = "Convert to WE format"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("720x540")
        self.minsize(620, 460)

        self.input_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self._output_dir_edited = False  # True once the user picks a folder themselves
        self.last_output = None
        self.msg_queue = queue.Queue()

        self._build_ui()
        self.after(100, self._poll_queue)

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        style = ttk.Style(self)
        style.configure("Title.TLabel", font=("Segoe UI", 13, "bold"))
        style.configure("Hint.TLabel", foreground="#555555")

        outer = ttk.Frame(self, padding=(20, 16, 20, 16))
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(4, weight=1)

        ttk.Label(
            outer, text="TestLink export  \u2192  WE / TelecomEgypt Word document",
            style="Title.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            outer,
            text="Removes the TestLink logo and branding and applies the WE header and footer.",
            style="Hint.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, 12))

        self._path_row(outer, 2, "TestLink export (.doc / .html)", self.input_path, self._browse_input)
        self._path_row(outer, 3, "Save the Word document in", self.output_dir, self._browse_output)

        # Log box (row 4 expands)
        log_frame = ttk.Frame(outer)
        log_frame.grid(row=4, column=0, sticky="nsew", pady=(12, 8))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log = tk.Text(log_frame, wrap="word", height=10, state="disabled",
                           relief="solid", borderwidth=1, font=("Consolas", 9))
        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        # Action row
        action = ttk.Frame(outer)
        action.grid(row=5, column=0, sticky="ew")
        action.columnconfigure(0, weight=1)
        self.run_btn = ttk.Button(action, text=RUN_TEXT, command=self._start_run)
        self.run_btn.grid(row=0, column=0, sticky="ew", ipady=6)
        self.show_btn = ttk.Button(action, text="Open output folder", state="disabled",
                                   command=self._show_in_folder)
        self.show_btn.grid(row=0, column=1, padx=(10, 0), ipady=6)

        self.progress = ttk.Progressbar(outer, mode="indeterminate")
        self.progress.grid(row=6, column=0, sticky="ew", pady=(10, 0))

        ttk.Label(self, text="\u00a9 Ahmed Saeed", foreground="#999999",
                 font=("Segoe UI", 8)).place(relx=1.0, rely=1.0, anchor="se", x=-8, y=-4)

        self._log_line("Pick a TestLink export to get started.")

    def _path_row(self, parent, row, label, variable, command):
        frame = ttk.Frame(parent)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 6))
        frame.columnconfigure(0, weight=1)
        ttk.Label(frame, text=label).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Entry(frame, textvariable=variable).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(frame, text="Browse...", command=command).grid(row=1, column=1)

    # ----------------------------------------------------------- callbacks
    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Choose the TestLink export",
            filetypes=[("TestLink export", "*.doc *.html *.htm"), ("All files", "*.*")],
        )
        if not path:
            return
        self.input_path.set(path)
        if not self._output_dir_edited:
            self.output_dir.set(os.path.dirname(path))

    def _browse_output(self):
        folder = filedialog.askdirectory(title="Choose where to save the Word document")
        if folder:
            self.output_dir.set(folder)
            self._output_dir_edited = True

    def _log_line(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_busy(self, busy):
        if busy:
            self.run_btn.configure(state="disabled", text="Working...")
            self.show_btn.configure(state="disabled")
            self.progress.start(12)
        else:
            self.progress.stop()
            self.run_btn.configure(state="normal", text=RUN_TEXT)

    def _start_run(self):
        src = self.input_path.get().strip()
        if not src:
            self._log_line("Pick a TestLink export file first.")
            return
        out_dir = self.output_dir.get().strip() or os.path.dirname(os.path.abspath(src))
        self.last_output = None
        self._log_line("-" * 40)
        self._set_busy(True)
        threading.Thread(target=self._worker, args=(src, out_dir), daemon=True).start()

    def _worker(self, src, out_dir):
        try:
            out = skill_core.run_conversion(
                src, out_dir, log=lambda m: self.msg_queue.put(("log", m))
            )
            self.msg_queue.put(("done", out))
        except skill_core.ConversionError as e:
            self.msg_queue.put(("error", str(e)))
        except Exception:
            # Unexpected problem: keep the details in the log so it can be reported.
            self.msg_queue.put(("crash", traceback.format_exc()))

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == "log":
                    self._log_line(payload)
                elif kind == "done":
                    self.last_output = payload
                    self._set_busy(False)
                    self.show_btn.configure(state="normal")
                    messagebox.showinfo(APP_TITLE, f"Done!\n\nSaved as:\n{os.path.basename(payload)}")
                elif kind == "error":
                    self._set_busy(False)
                    self._log_line("ERROR: " + payload)
                    messagebox.showerror(APP_TITLE, payload)
                elif kind == "crash":
                    self._set_busy(False)
                    self._log_line("Something unexpected went wrong:\n" + payload)
                    messagebox.showerror(
                        APP_TITLE,
                        "Something unexpected went wrong. The details are in the log "
                        "box, so you can copy them and send them to whoever maintains this tool.",
                    )
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _show_in_folder(self):
        path = self.last_output
        if not path or not os.path.exists(path):
            return
        folder = os.path.dirname(path)
        if hasattr(os, "startfile"):      # Windows: open the folder in Explorer
            os.startfile(folder)
        else:                             # other systems (development only)
            self._log_line(f"Saved in: {folder}")


if __name__ == "__main__":
    App().mainloop()
