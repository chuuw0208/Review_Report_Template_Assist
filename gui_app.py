"""
Phase 3 – Desktop GUI Application for MRM Report Content Migration.

A lightweight, native Tkinter desktop interface that enables model risk
validators to select a Template and a Reference Report, preview detected
report type, run the migration pipeline with live streamed logs, and
open the generated draft document directly in Word.

Usage:
    python gui_app.py
"""

import os
import sys
import queue
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# Ensure the local directory is on sys.path so phase2_core is importable
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

try:
    from phase2_core import run_migration, detect_report_type
except ImportError:
    run_migration = None
    detect_report_type = None

# Enable Windows DPI awareness for sharp text rendering on modern displays
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass


# ============================================================================
# THREAD-SAFE STDOUT REDIRECTOR
# ============================================================================

class ThreadSafeLogStream:
    """Redirects writes to a queue so the Tkinter mainloop can update the UI safely."""
    def __init__(self, log_queue, original_stream):
        self.log_queue = log_queue
        self.original_stream = original_stream

    def write(self, text):
        if text:
            self.log_queue.put(text)
        if self.original_stream:
            self.original_stream.write(text)
            self.original_stream.flush()

    def flush(self):
        if self.original_stream:
            self.original_stream.flush()


# ============================================================================
# GUI APPLICATION
# ============================================================================

class MRMAutomationApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Model Risk Management — Report Drafting Tool")
        self.geometry("820x680")
        self.minsize(720, 560)

        # Application state
        self.log_queue = queue.Queue()
        self.is_processing = False
        self.generated_output_path = None

        self._configure_styles()
        self._build_ui()

        # Start background queue poller for log streaming
        self.after(100, self._process_log_queue)

    def _configure_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # Color palette
        self.c_bg = "#f5f6f8"
        self.c_card = "#ffffff"
        self.c_primary = "#0b6623"     # Bank dark green accent
        self.c_primary_active = "#084d1a"
        self.c_border = "#d0d4dc"
        self.c_text = "#1f242d"
        self.c_subtext = "#5a6270"

        self.configure(bg=self.c_bg)

        style.configure("TFrame", background=self.c_bg)
        style.configure("Card.TFrame", background=self.c_card, relief="solid", borderwidth=1)
        style.configure("TLabel", background=self.c_bg, foreground=self.c_text, font=("Segoe UI", 9))
        style.configure("Card.TLabel", background=self.c_card, foreground=self.c_text, font=("Segoe UI", 9))
        style.configure("Sub.TLabel", background=self.c_card, foreground=self.c_subtext, font=("Segoe UI", 8))
        style.configure("Bold.TLabel", background=self.c_card, foreground=self.c_text, font=("Segoe UI", 9, "bold"))

        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            foreground="#ffffff",
            background=self.c_primary,
            padding=(16, 8)
        )
        style.map(
            "Primary.TButton",
            background=[("active", self.c_primary_active), ("disabled", "#a0b5a6")]
        )

        style.configure("Secondary.TButton", font=("Segoe UI", 9), padding=(10, 5))

    def _build_ui(self):
        container = ttk.Frame(self, padding="16 16 16 16")
        container.pack(fill=tk.BOTH, expand=True)

        # ---------------- 1. Header ----------------
        header_frame = ttk.Frame(container)
        header_frame.pack(fill=tk.X, pady=(0, 14))

        title_lbl = tk.Label(
            header_frame,
            text="MRM Report Content Migration Tool",
            font=("Segoe UI", 14, "bold"),
            bg=self.c_bg,
            fg="#0b6623"
        )
        title_lbl.pack(anchor="w")

        subtitle_lbl = tk.Label(
            header_frame,
            text="Intelligently migrate section contents from a finalized Reference Report into a New Template with native Track Changes.",
            font=("Segoe UI", 9),
            bg=self.c_bg,
            fg=self.c_subtext
        )
        subtitle_lbl.pack(anchor="w", pady=(2, 0))

        # ---------------- 2. File Selection Card ----------------
        files_card = ttk.Frame(container, style="Card.TFrame", padding="14 12 14 12")
        files_card.pack(fill=tk.X, pady=(0, 12))

        # Row 1: New Template
        lbl_template = ttk.Label(files_card, text="1. New Report Template (.docx):", style="Bold.TLabel")
        lbl_template.grid(row=0, column=0, sticky="w", pady=(2, 2))

        self.var_template = tk.StringVar()
        self.ent_template = ttk.Entry(files_card, textvariable=self.var_template, font=("Segoe UI", 9))
        self.ent_template.grid(row=1, column=0, sticky="ew", padx=(0, 8), pady=(0, 10))

        btn_browse_template = ttk.Button(
            files_card, text="Browse...", style="Secondary.TButton", command=self._browse_template
        )
        btn_browse_template.grid(row=1, column=1, sticky="e", pady=(0, 10))

        # Row 2: Reference Report
        lbl_ref = ttk.Label(files_card, text="2. Reference Report (.docx):", style="Bold.TLabel")
        lbl_ref.grid(row=2, column=0, sticky="w", pady=(2, 2))

        self.var_reference = tk.StringVar()
        self.ent_reference = ttk.Entry(files_card, textvariable=self.var_reference, font=("Segoe UI", 9))
        self.ent_reference.grid(row=3, column=0, sticky="ew", padx=(0, 8), pady=(0, 4))

        btn_browse_ref = ttk.Button(
            files_card, text="Browse...", style="Secondary.TButton", command=self._browse_reference
        )
        btn_browse_ref.grid(row=3, column=1, sticky="e", pady=(0, 4))

        # Badge: Detected Report Type
        self.lbl_report_type = ttk.Label(
            files_card,
            text="Detected Type: (Select reference file to auto-detect)",
            style="Sub.TLabel"
        )
        self.lbl_report_type.grid(row=4, column=0, sticky="w", pady=(0, 10))

        # Row 3: Output Destination
        lbl_out = ttk.Label(files_card, text="3. Output Draft Report (.docx):", style="Bold.TLabel")
        lbl_out.grid(row=5, column=0, sticky="w", pady=(2, 2))

        self.var_output = tk.StringVar()
        self.ent_output = ttk.Entry(files_card, textvariable=self.var_output, font=("Segoe UI", 9))
        self.ent_output.grid(row=6, column=0, sticky="ew", padx=(0, 8), pady=(0, 2))

        btn_browse_out = ttk.Button(
            files_card, text="Browse...", style="Secondary.TButton", command=self._browse_output
        )
        btn_browse_out.grid(row=6, column=1, sticky="e", pady=(0, 2))

        files_card.columnconfigure(0, weight=1)

        # ---------------- 3. Action & Status Bar ----------------
        action_bar = ttk.Frame(container)
        action_bar.pack(fill=tk.X, pady=(0, 10))

        self.btn_run = ttk.Button(
            action_bar,
            text="▶  Generate Draft Report",
            style="Primary.TButton",
            command=self._start_pipeline
        )
        self.btn_run.pack(side=tk.LEFT, padx=(0, 12))

        self.progress_bar = ttk.Progressbar(action_bar, mode="indeterminate", length=180)
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 12))

        self.lbl_status = ttk.Label(action_bar, text="Ready", font=("Segoe UI", 9, "italic"))
        self.lbl_status.pack(side=tk.RIGHT)

        # ---------------- 4. Live Log Console ----------------
        console_card = ttk.Frame(container, style="Card.TFrame", padding="10 10 10 10")
        console_card.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        console_header = ttk.Frame(console_card, style="Card.TFrame")
        console_header.pack(fill=tk.X, pady=(0, 6))

        lbl_console_title = ttk.Label(console_header, text="Execution Log:", style="Bold.TLabel")
        lbl_console_title.pack(side=tk.LEFT)

        btn_clear_log = ttk.Button(
            console_header, text="Clear", style="Secondary.TButton", command=self._clear_log
        )
        btn_clear_log.pack(side=tk.RIGHT)

        self.txt_log = scrolledtext.ScrolledText(
            console_card,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg="#1e1e1e",
            fg="#d4d4d4",
            insertbackground="#ffffff",
            relief="flat",
            borderwidth=0
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True)

        # Configure highlighting tags for log clarity
        self.txt_log.tag_config("INFO", foreground="#70c0ff")
        self.txt_log.tag_config("SUCCESS", foreground="#68e590", font=("Consolas", 9, "bold"))
        self.txt_log.tag_config("DELETE", foreground="#f27d7d")
        self.txt_log.tag_config("MIGRATE", foreground="#ffd370")
        self.txt_log.tag_config("ERROR", foreground="#ff4d4f", font=("Consolas", 9, "bold"))

        # ---------------- 5. Footer Quick Actions ----------------
        footer_frame = ttk.Frame(container)
        footer_frame.pack(fill=tk.X)

        self.btn_open_doc = ttk.Button(
            footer_frame,
            text="📄  Open Draft Document in Word",
            style="Secondary.TButton",
            state=tk.DISABLED,
            command=self._open_output_document
        )
        self.btn_open_doc.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_open_dir = ttk.Button(
            footer_frame,
            text="📁  Open Output Folder",
            style="Secondary.TButton",
            state=tk.DISABLED,
            command=self._open_output_folder
        )
        self.btn_open_dir.pack(side=tk.LEFT)

    # ========================================================================
    # FILE BROWSING & AUTO-DETECTION
    # ========================================================================

    def _browse_template(self):
        filename = filedialog.askopenfilename(
            title="Select New Report Template",
            filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")]
        )
        if filename:
            self.var_template.set(filename)
            # Auto-suggest output path if not already customized
            if not self.var_output.get():
                base, ext = os.path.splitext(filename)
                self.var_output.set(f"{base}_Draft{ext}")

    def _browse_reference(self):
        filename = filedialog.askopenfilename(
            title="Select Reference Report",
            filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")]
        )
        if filename:
            self.var_reference.set(filename)
            self._update_detected_type(filename)

    def _browse_output(self):
        filename = filedialog.asksaveasfilename(
            title="Choose Output Draft Path",
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")]
        )
        if filename:
            self.var_output.set(filename)

    def _update_detected_type(self, filepath):
        if detect_report_type is None:
            return
        try:
            rtype = detect_report_type(filepath)
            self.lbl_report_type.config(
                text=f"Detected Type:  ✔ {rtype} Report",
                foreground="#0b6623"
            )
        except Exception:
            self.lbl_report_type.config(
                text="Detected Type:  ⚠ Keyword not found (ensure Assessment / Affirmation / Validation in name)",
                foreground="#b07000"
            )

    # ========================================================================
    # PIPELINE EXECUTION (BACKGROUND THREAD)
    # ========================================================================

    def _start_pipeline(self):
        if self.is_processing:
            return

        template_path = self.var_template.get().strip()
        reference_path = self.var_reference.get().strip()
        output_path = self.var_output.get().strip()

        # Validation
        if not template_path:
            messagebox.showwarning("Missing Input", "Please select a New Report Template (.docx).")
            return
        if not os.path.isfile(template_path):
            messagebox.showerror("File Error", f"Template file does not exist:\n{template_path}")
            return

        if not reference_path:
            messagebox.showwarning("Missing Input", "Please select a Reference Report (.docx).")
            return
        if not os.path.isfile(reference_path):
            messagebox.showerror("File Error", f"Reference report does not exist:\n{reference_path}")
            return

        if not output_path:
            base, ext = os.path.splitext(template_path)
            output_path = f"{base}_Draft{ext}"
            self.var_output.set(output_path)

        if os.path.abspath(template_path) == os.path.abspath(output_path):
            messagebox.showerror(
                "Invalid Output Path",
                "Output path cannot be the same file as the template!\n"
                "Please choose a distinct output name to protect the template."
            )
            return

        # Prepare UI for processing
        self.is_processing = True
        self.btn_run.config(state=tk.DISABLED)
        self.btn_open_doc.config(state=tk.DISABLED)
        self.btn_open_dir.config(state=tk.DISABLED)
        self.progress_bar.start(10)
        self.lbl_status.config(text="Processing Word documents in background...")
        self._append_log("\n" + "=" * 60 + "\n[START] Starting Report Draft Generation...\n" + "=" * 60 + "\n")

        # Launch worker thread
        worker = threading.Thread(
            target=self._run_migration_thread,
            args=(template_path, reference_path, output_path),
            daemon=True
        )
        worker.start()

    def _run_migration_thread(self, template_path, reference_path, output_path):
        """Worker thread executing the COM migration pipeline with stdout redirected."""
        original_stdout = sys.stdout
        original_stderr = sys.stderr

        sys.stdout = ThreadSafeLogStream(self.log_queue, original_stdout)
        sys.stderr = ThreadSafeLogStream(self.log_queue, original_stderr)

        success = False
        error_message = None

        try:
            if run_migration is None:
                raise RuntimeError("phase2_core module could not be loaded.")

            run_migration(template_path, reference_path, output_path)
            success = True
            self.generated_output_path = output_path
        except Exception as exc:
            error_message = str(exc)
        finally:
            sys.stdout = original_stdout
            sys.stderr = original_stderr
            self.after(0, self._on_pipeline_completed, success, error_message, output_path)

    def _on_pipeline_completed(self, success, error_message, output_path):
        """Called on main UI thread once migration worker finishes."""
        self.is_processing = False
        self.progress_bar.stop()
        self.btn_run.config(state=tk.NORMAL)

        if success:
            self.lbl_status.config(text="✔  Draft generated successfully!")
            self.btn_open_doc.config(state=tk.NORMAL)
            self.btn_open_dir.config(state=tk.NORMAL)

            messagebox.showinfo(
                "Migration Complete",
                f"Draft report successfully created!\n\n"
                f"Saved to:\n{output_path}\n\n"
                f"You can now open the document in Word to review the tracked revisions."
            )
        else:
            self.lbl_status.config(text="✖  Process encountered an error.")
            self._append_log(f"\n[ERROR] {error_message}\n")
            messagebox.showerror(
                "Execution Failed",
                f"An error occurred while generating the draft:\n\n{error_message}\n\n"
                f"Please inspect the log window for details."
            )

    # ========================================================================
    # LOGGING & QUICK ACTIONS
    # ========================================================================

    def _process_log_queue(self):
        """Poll the log queue and flush text into the ScrolledText widget."""
        while not self.log_queue.empty():
            try:
                msg = self.log_queue.get_nowait()
                self._append_log(msg)
            except queue.Empty:
                break
        self.after(100, self._process_log_queue)

    def _append_log(self, text):
        self.txt_log.insert(tk.END, text)

        # Apply coloring based on tags
        for line in text.splitlines():
            if "[INFO]" in line:
                pass
            elif "[DELETE]" in line:
                pass
            elif "[MIGRATE]" in line:
                pass
            elif "ERROR" in line:
                pass

        self.txt_log.see(tk.END)

    def _clear_log(self):
        self.txt_log.delete("1.0", tk.END)

    def _open_output_document(self):
        if self.generated_output_path and os.path.exists(self.generated_output_path):
            try:
                os.startfile(self.generated_output_path)
            except AttributeError:
                # macOS / Linux fallback for local development testing
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", self.generated_output_path])
            except Exception as exc:
                messagebox.showerror("Error", f"Failed to open file:\n{exc}")

    def _open_output_folder(self):
        if self.generated_output_path and os.path.exists(self.generated_output_path):
            folder = os.path.dirname(os.path.abspath(self.generated_output_path))
            try:
                if sys.platform == "win32":
                    subprocess.Popen(f'explorer /select,"{os.path.abspath(self.generated_output_path)}"')
                else:
                    subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", folder])
            except Exception as exc:
                messagebox.showerror("Error", f"Failed to open folder:\n{exc}")


# ============================================================================
# ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    app = MRMAutomationApp()
    app.mainloop()

