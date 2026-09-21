"""TestLink -> WE/TelecomEgypt formatter: desktop GUI.

All real work lives in skill_core.py and runs on a background thread so the
window never freezes. Progress and errors come back through a queue that the
main thread drains with .after() (widgets are only touched from the main thread).
"""
import os
import queue
import subprocess
import sys
import threading
import traceback
from tkinter import filedialog, messagebox

import customtkinter as ctk

import skill_core

APP_TITLE = "TestLink to WE Formatter"
RUN_TEXT = "Convert to WE format"


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("720x560")
        self.minsize(620, 480)
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.input_path = ctk.StringVar()
        self.output_dir = ctk.StringVar()
        self._output_dir_edited = False  # True once the user picks a folder themselves
        self.last_output = None
        self.msg_queue = queue.Queue()

        self._build_ui()
        self.after(100, self._poll_queue)

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1)

        ctk.CTkLabel(
            self, text="TestLink export  \u2192  WE / TelecomEgypt Word document",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=20, pady=(18, 2))
        ctk.CTkLabel(
            self,
            text="Removes the TestLink logo and branding and applies the WE header and footer.",
            text_color=("gray35", "gray70"),
        ).grid(row=1, column=0, sticky="w", padx=20, pady=(0, 12))

        # Input file row
        self._path_row(
            row=2, label="TestLink export (.doc / .html)",
            variable=self.input_path, command=self._browse_input,
        )
        # Output folder row
        self._path_row(
            row=3, label="Save the Word document in",
            variable=self.output_dir, command=self._browse_output,
        )

        # Action row
        action = ctk.CTkFrame(self, fg_color="transparent")
        action.grid(row=4, column=0, sticky="ew", padx=20, pady=(8, 4))
        action.grid_columnconfigure(0, weight=1)
        self.run_btn = ctk.CTkButton(
            action, text=RUN_TEXT, height=38, command=self._start_run,
        )
        self.run_btn.grid(row=0, column=0, sticky="ew")
        self.show_btn = ctk.CTkButton(
            action, text="Show in folder", height=38, width=140,
            fg_color=("gray70", "gray30"), hover_color=("gray60", "gray40"),
            state="disabled", command=self._show_in_folder,
        )
        self.show_btn.grid(row=0, column=1, padx=(10, 0))

        # Log box
        self.log = ctk.CTkTextbox(self, wrap="word")
        self.log.grid(row=5, column=0, sticky="nsew", padx=20, pady=(8, 8))
        self.log.configure(state="disabled")

        self.progress = ctk.CTkProgressBar(self, mode="indeterminate")
        self.progress.grid(row=6, column=0, sticky="ew", padx=20, pady=(0, 18))
        self.progress.set(0)

        self._log_line("Pick a TestLink export to get started.")

    def _path_row(self, row, label, variable, command):
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.grid(row=row, column=0, sticky="ew", padx=20, pady=(4, 4))
        frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(frame, text=label, anchor="w").grid(row=0, column=0, columnspan=2, sticky="w")
        ctk.CTkEntry(frame, textvariable=variable).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkButton(frame, text="Browse...", width=100, command=command).grid(row=1, column=1)

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
            self.progress.start()
        else:
            self.progress.stop()
            self.progress.set(0)
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
        if sys.platform.startswith("win"):
            subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", "-R", path])
        else:
            subprocess.Popen(["xdg-open", os.path.dirname(path)])


if __name__ == "__main__":
    App().mainloop()
