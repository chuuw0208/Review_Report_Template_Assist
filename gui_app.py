"""
MRMV Report Content Migration Tool — Desktop GUI
Refined modern interface tailored for Citizens Bank Model Risk Management & Validation (MRMV).
Features:
  - Apple / Microsoft Teams clean minimalist design with Citizens Bank Green (#008450)
  - Side-by-side (左右) upload layout for New Report Template and Reference Report
  - Dual upload support: Drag & Drop (.docx) AND Click to Browse
  - Target Report Type segmented toggle: Assessment vs Affirmation
  - Reference Report auto-detection badge
  - Automated background output naming (no manual output upload needed)
  - Real-time streamed log console and one-click launch in MS Word
"""

import os
import sys
import queue
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# Ensure local script directory is on sys.path
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

try:
    from phase2_core import run_migration, detect_report_type
except ImportError:
    run_migration = None
    detect_report_type = None

# Enable Windows High-DPI Awareness for crisp rendering
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# Drag and Drop support via tkinterdnd2 (graceful fallback if not installed)
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    BaseWindow = TkinterDnD.Tk
    HAS_DND = True
except Exception:
    BaseWindow = tk.Tk
    HAS_DND = False
    DND_FILES = None


# ============================================================================
# THREAD-SAFE STDOUT STREAM REDIRECTOR
# ============================================================================

class ThreadSafeLogStream:
    """Intercepts stdout/stderr and feeds a thread-safe Queue for UI rendering."""
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
# MODERN DESKTOP GUI
# ============================================================================

class MRMAutomationApp(BaseWindow):
    def __init__(self):
        super().__init__()

        self.title("MRMV Report Content Migration Tool")
        self.geometry("960x780")
        self.minsize(880, 680)

        # Citizens Bank Brand Palette
        self.c_brand = "#008450"          # Citizens Primary Green
        self.c_brand_hover = "#00663d"    # Citizens Dark Green
        self.c_brand_light = "#e8f5ed"    # Citizens Soft Mint Tint
        self.c_brand_border = "#b9e2cb"   # Mint border

        # Neutral Slate Canvas & Card Colors
        self.c_canvas = "#f8fafc"         # Clean, soft background
        self.c_card = "#ffffff"           # Pure card white
        self.c_border = "#e2e8f0"         # Crisp card border
        self.c_drop_bg = "#fbfcfd"        # Dropzone soft fill
        self.c_drop_hover = "#f0faf4"     # Dropzone hover fill
        self.c_drop_border = "#cbd5e1"    # Dropzone subtle border
        self.c_text_main = "#0f172a"      # High-contrast slate heading
        self.c_text_sub = "#475569"       # Refined subtext
        self.c_text_muted = "#94a3b8"     # Inactive/muted text

        self.configure(bg=self.c_canvas)

        # State variables
        self.log_queue = queue.Queue()
        self.is_processing = False
        self.template_path = None
        self.reference_path = None
        self.generated_output_path = None
        self.target_report_type = "Assessment"  # Default: Assessment

        self._configure_styles()
        self._build_ui()

        # Start log consumer loop
        self.after(80, self._process_log_queue)

    def _configure_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("Canvas.TFrame", background=self.c_canvas)
        style.configure("Card.TFrame", background=self.c_card, relief="flat")
        style.configure("Brand.Horizontal.TProgressbar", troughcolor="#e2e8f0", background=self.c_brand)

    def _build_ui(self):
        # Outer padding wrapper
        root_padding = tk.Frame(self, bg=self.c_canvas, padx=28, pady=24)
        root_padding.pack(fill=tk.BOTH, expand=True)

        # ---------------- 1. Top Header ----------------
        header_frame = tk.Frame(root_padding, bg=self.c_canvas)
        header_frame.pack(fill=tk.X, pady=(0, 20))

        # Brand pill / tag
        brand_tag_frame = tk.Frame(header_frame, bg=self.c_canvas)
        brand_tag_frame.pack(anchor="w", pady=(0, 4))

        brand_dot = tk.Label(brand_tag_frame, text="●", font=("Segoe UI", 9, "bold"), fg=self.c_brand, bg=self.c_canvas)
        brand_dot.pack(side=tk.LEFT, padx=(0, 6))

        brand_tag = tk.Label(
            brand_tag_frame,
            text="CITIZENS MODEL RISK MANAGEMENT & VALIDATION",
            font=("Segoe UI", 8, "bold"),
            fg=self.c_brand,
            bg=self.c_canvas
        )
        brand_tag.pack(side=tk.LEFT)

        # Main Title (exact text requested)
        title_label = tk.Label(
            header_frame,
            text="MRMV Report Content Migration Tool",
            font=("Segoe UI", 18, "bold"),
            fg=self.c_text_main,
            bg=self.c_canvas
        )
        title_label.pack(anchor="w")

        # ---------------- 2. Side-by-Side Upload Cards (左右排版) ----------------
        cards_container = tk.Frame(root_padding, bg=self.c_canvas)
        cards_container.pack(fill=tk.X, pady=(0, 18))
        cards_container.columnconfigure(0, weight=1, uniform="group1")
        cards_container.columnconfigure(1, weight=1, uniform="group1")

        # --- LEFT CARD: New Report Template ---
        self.left_card = tk.Frame(
            cards_container,
            bg=self.c_card,
            highlightbackground=self.c_border,
            highlightthickness=1,
            padx=18,
            pady=16
        )
        self.left_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        left_header = tk.Frame(self.left_card, bg=self.c_card)
        left_header.pack(fill=tk.X, pady=(0, 10))

        left_title = tk.Label(
            left_header,
            text="1. New Report Template",
            font=("Segoe UI", 11, "bold"),
            fg=self.c_text_main,
            bg=self.c_card
        )
        left_title.pack(side=tk.LEFT)

        # Left Dropzone
        self.template_dropzone = tk.Frame(
            self.left_card,
            bg=self.c_drop_bg,
            highlightbackground=self.c_drop_border,
            highlightthickness=1,
            padx=16,
            pady=22,
            cursor="hand2"
        )
        self.template_dropzone.pack(fill=tk.X, pady=(0, 14))

        self.lbl_tmpl_icon = tk.Label(
            self.template_dropzone,
            text="📄",
            font=("Segoe UI Emoji", 26),
            bg=self.c_drop_bg,
            fg=self.c_brand
        )
        self.lbl_tmpl_icon.pack(pady=(0, 4))

        self.lbl_tmpl_main = tk.Label(
            self.template_dropzone,
            text="Drag & drop template here",
            font=("Segoe UI", 10, "bold"),
            fg=self.c_text_main,
            bg=self.c_drop_bg
        )
        self.lbl_tmpl_main.pack()

        self.lbl_tmpl_sub = tk.Label(
            self.template_dropzone,
            text="or click to browse .docx file",
            font=("Segoe UI", 8),
            fg=self.c_text_sub,
            bg=self.c_drop_bg
        )
        self.lbl_tmpl_sub.pack(pady=(2, 6))

        self.btn_tmpl_browse = tk.Label(
            self.template_dropzone,
            text="Browse Template",
            font=("Segoe UI", 9, "bold"),
            fg=self.c_brand,
            bg=self.c_brand_light,
            padx=12,
            pady=4,
            cursor="hand2"
        )
        self.btn_tmpl_browse.pack()

        # Bind clicks & hover for template dropzone
        self._bind_dropzone_events(self.template_dropzone, self._browse_template)

        # Target Report Type Toggle Section
        target_section = tk.Frame(self.left_card, bg=self.c_card)
        target_section.pack(fill=tk.X, pady=(4, 2))

        target_label = tk.Label(
            target_section,
            text="Desired Output Report Type:",
            font=("Segoe UI", 9, "bold"),
            fg=self.c_text_main,
            bg=self.c_card
        )
        target_label.pack(anchor="w", pady=(0, 6))

        # Apple-style Segmented Pill Control
        segmented_frame = tk.Frame(target_section, bg="#e2e8f0", padx=3, pady=3)
        segmented_frame.pack(fill=tk.X)

        self.btn_pill_assessment = tk.Label(
            segmented_frame,
            text="Assessment Report",
            font=("Segoe UI", 9, "bold"),
            bg=self.c_brand,
            fg="#ffffff",
            padx=16,
            pady=6,
            cursor="hand2"
        )
        self.btn_pill_assessment.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))

        self.btn_pill_affirmation = tk.Label(
            segmented_frame,
            text="Affirmation Report",
            font=("Segoe UI", 9, "bold"),
            bg="#ffffff",
            fg=self.c_text_sub,
            padx=16,
            pady=6,
            cursor="hand2"
        )
        self.btn_pill_affirmation.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn_pill_assessment.bind("<Button-1>", lambda e: self._select_report_type("Assessment"))
        self.btn_pill_affirmation.bind("<Button-1>", lambda e: self._select_report_type("Affirmation"))

        # --- RIGHT CARD: Reference Report ---
        self.right_card = tk.Frame(
            cards_container,
            bg=self.c_card,
            highlightbackground=self.c_border,
            highlightthickness=1,
            padx=18,
            pady=16
        )
        self.right_card.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        right_header = tk.Frame(self.right_card, bg=self.c_card)
        right_header.pack(fill=tk.X, pady=(0, 10))

        right_title = tk.Label(
            right_header,
            text="2. Reference Report",
            font=("Segoe UI", 11, "bold"),
            fg=self.c_text_main,
            bg=self.c_card
        )
        right_title.pack(side=tk.LEFT)

        # Right Dropzone
        self.ref_dropzone = tk.Frame(
            self.right_card,
            bg=self.c_drop_bg,
            highlightbackground=self.c_drop_border,
            highlightthickness=1,
            padx=16,
            pady=22,
            cursor="hand2"
        )
        self.ref_dropzone.pack(fill=tk.X, pady=(0, 14))

        self.lbl_ref_icon = tk.Label(
            self.ref_dropzone,
            text="📑",
            font=("Segoe UI Emoji", 26),
            bg=self.c_drop_bg,
            fg="#0284c7"
        )
        self.lbl_ref_icon.pack(pady=(0, 4))

        self.lbl_ref_main = tk.Label(
            self.ref_dropzone,
            text="Drag & drop reference here",
            font=("Segoe UI", 10, "bold"),
            fg=self.c_text_main,
            bg=self.c_drop_bg
        )
        self.lbl_ref_main.pack()

        self.lbl_ref_sub = tk.Label(
            self.ref_dropzone,
            text="or click to browse .docx file",
            font=("Segoe UI", 8),
            fg=self.c_text_sub,
            bg=self.c_drop_bg
        )
        self.lbl_ref_sub.pack(pady=(2, 6))

        self.btn_ref_browse = tk.Label(
            self.ref_dropzone,
            text="Browse Reference",
            font=("Segoe UI", 9, "bold"),
            fg="#0369a1",
            bg="#e0f2fe",
            padx=12,
            pady=4,
            cursor="hand2"
        )
        self.btn_ref_browse.pack()

        # Bind clicks & hover for reference dropzone
        self._bind_dropzone_events(self.ref_dropzone, self._browse_reference)

        # Auto-Detection Badge & Concise Explanatory Note
        ref_status_section = tk.Frame(self.right_card, bg=self.c_card)
        ref_status_section.pack(fill=tk.X, pady=(4, 2))

        ref_status_title = tk.Label(
            ref_status_section,
            text="Reference Classification:",
            font=("Segoe UI", 9, "bold"),
            fg=self.c_text_main,
            bg=self.c_card
        )
        ref_status_title.pack(anchor="w", pady=(0, 6))

        self.badge_ref_type = tk.Label(
            ref_status_section,
            text="Type: Auto-detected from filename",
            font=("Segoe UI", 9),
            bg="#f1f5f9",
            fg=self.c_text_sub,
            padx=12,
            pady=7
        )
        self.badge_ref_type.pack(fill=tk.X)

        # Setup OS Drag-and-Drop if TkinterDnD is available
        self._setup_drag_and_drop()

        # ---------------- 3. Central Action & Execution Card ----------------
        action_card = tk.Frame(
            root_padding,
            bg=self.c_card,
            highlightbackground=self.c_border,
            highlightthickness=1,
            padx=18,
            pady=14
        )
        action_card.pack(fill=tk.X, pady=(0, 16))

        # Main Action Button (Apple / Citizens pill style)
        btn_container = tk.Frame(action_card, bg=self.c_card)
        btn_container.pack(fill=tk.X)

        self.btn_generate = tk.Button(
            btn_container,
            text="Generate Draft Report",
            font=("Segoe UI", 11, "bold"),
            bg=self.c_brand,
            fg="#ffffff",
            activebackground=self.c_brand_hover,
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            padx=28,
            pady=10,
            cursor="hand2",
            command=self._start_pipeline
        )
        self.btn_generate.pack(side=tk.LEFT)

        # Progress bar
        self.progress_bar = ttk.Progressbar(
            btn_container,
            style="Brand.Horizontal.TProgressbar",
            mode="indeterminate",
            length=220
        )
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(18, 18))

        # Status text
        self.lbl_status = tk.Label(
            btn_container,
            text="Ready",
            font=("Segoe UI", 9, "italic"),
            fg=self.c_text_sub,
            bg=self.c_card
        )
        self.lbl_status.pack(side=tk.RIGHT)

        # ---------------- 4. Glass-Style Live Activity Log ----------------
        log_card = tk.Frame(
            root_padding,
            bg="#111827",
            highlightbackground="#1f2937",
            highlightthickness=1,
            padx=14,
            pady=12
        )
        log_card.pack(fill=tk.BOTH, expand=True, pady=(0, 14))

        log_top_bar = tk.Frame(log_card, bg="#111827")
        log_top_bar.pack(fill=tk.X, pady=(0, 6))

        lbl_log_title = tk.Label(
            log_top_bar,
            text="ACTIVITY & COMPLIANCE LOG",
            font=("Consolas", 8, "bold"),
            fg="#9ca3af",
            bg="#111827"
        )
        lbl_log_title.pack(side=tk.LEFT)

        btn_clear = tk.Label(
            log_top_bar,
            text="Clear",
            font=("Segoe UI", 8),
            fg="#9ca3af",
            bg="#1f2937",
            padx=8,
            pady=2,
            cursor="hand2"
        )
        btn_clear.pack(side=tk.RIGHT)
        btn_clear.bind("<Button-1>", lambda e: self._clear_log())

        self.txt_log = scrolledtext.ScrolledText(
            log_card,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg="#111827",
            fg="#e5e7eb",
            insertbackground="#ffffff",
            relief="flat",
            bd=0
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True)

        # ---------------- 5. Footer Quick Actions ----------------
        footer = tk.Frame(root_padding, bg=self.c_canvas)
        footer.pack(fill=tk.X)

        self.btn_open_word = tk.Label(
            footer,
            text="📄  Open Draft Document in Word",
            font=("Segoe UI", 9, "bold"),
            bg="#e2e8f0",
            fg=self.c_text_muted,
            padx=16,
            pady=7
        )
        self.btn_open_word.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_open_folder = tk.Label(
            footer,
            text="📁  Open Containing Folder",
            font=("Segoe UI", 9, "bold"),
            bg="#e2e8f0",
            fg=self.c_text_muted,
            padx=16,
            pady=7
        )
        self.btn_open_folder.pack(side=tk.LEFT)

    # ========================================================================
    # INTERACTION & DRAG-AND-DROP SETUP
    # ========================================================================

    def _bind_dropzone_events(self, dropzone_widget, browse_command):
        """Bind click and hover interactions to a dropzone and all its children."""
        def on_click(e):
            browse_command()

        def on_enter(e):
            dropzone_widget.configure(bg=self.c_drop_hover)
            for child in dropzone_widget.winfo_children():
                if child.cget("text") not in ("Browse Template", "Browse Reference"):
                    child.configure(bg=self.c_drop_hover)

        def on_leave(e):
            dropzone_widget.configure(bg=self.c_drop_bg)
            for child in dropzone_widget.winfo_children():
                if child.cget("text") not in ("Browse Template", "Browse Reference"):
                    child.configure(bg=self.c_drop_bg)

        dropzone_widget.bind("<Button-1>", on_click)
        dropzone_widget.bind("<Enter>", on_enter)
        dropzone_widget.bind("<Leave>", on_leave)

        for child in dropzone_widget.winfo_children():
            child.bind("<Button-1>", on_click)

    def _setup_drag_and_drop(self):
        """Enable OS drag-and-drop if TkinterDnD is available."""
        if not HAS_DND:
            return

        try:
            self.template_dropzone.drop_target_register(DND_FILES)
            self.template_dropzone.dnd_bind("<<Drop>>", self._on_drop_template)

            self.ref_dropzone.drop_target_register(DND_FILES)
            self.ref_dropzone.dnd_bind("<<Drop>>", self._on_drop_reference)
        except Exception:
            pass

    def _clean_drop_path(self, raw_data):
        """Clean dropped path from TkinterDnD (handles spaces and braces)."""
        if not raw_data:
            return None
        path = raw_data.strip()
        if path.startswith("{") and path.endswith("}"):
            path = path[1:-1]
        return path

    def _on_drop_template(self, event):
        path = self._clean_drop_path(event.data)
        if path and path.lower().endswith(".docx") and os.path.isfile(path):
            self._set_template_file(path)
        else:
            messagebox.showwarning("Invalid File", "Please drop a valid Word Document (.docx).")

    def _on_drop_reference(self, event):
        path = self._clean_drop_path(event.data)
        if path and path.lower().endswith(".docx") and os.path.isfile(path):
            self._set_reference_file(path)
        else:
            messagebox.showwarning("Invalid File", "Please drop a valid Word Document (.docx).")

    # ========================================================================
    # FILE SELECTION & STATE UPDATES
    # ========================================================================

    def _browse_template(self):
        path = filedialog.askopenfilename(
            title="Select New Report Template",
            filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")]
        )
        if path:
            self._set_template_file(path)

    def _set_template_file(self, path):
        self.template_path = os.path.abspath(path)
        basename = os.path.basename(self.template_path)
        size_kb = os.path.getsize(self.template_path) // 1024

        # Update Dropzone appearance to show file is loaded
        self.lbl_tmpl_icon.config(text="✔", fg=self.c_brand)
        self.lbl_tmpl_main.config(text=basename[:32] + ("..." if len(basename) > 32 else ""))
        self.lbl_tmpl_sub.config(text=f"{size_kb} KB  •  Click to change")
        self.btn_tmpl_browse.config(text="Change File", bg="#d1fae5", fg=self.c_brand)

        self._append_log(f"[INFO] Template selected: {basename}\n")

    def _browse_reference(self):
        path = filedialog.askopenfilename(
            title="Select Reference Report",
            filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")]
        )
        if path:
            self._set_reference_file(path)

    def _set_reference_file(self, path):
        self.reference_path = os.path.abspath(path)
        basename = os.path.basename(self.reference_path)
        size_kb = os.path.getsize(self.reference_path) // 1024

        self.lbl_ref_icon.config(text="✔", fg="#0284c7")
        self.lbl_ref_main.config(text=basename[:32] + ("..." if len(basename) > 32 else ""))
        self.lbl_ref_sub.config(text=f"{size_kb} KB  •  Click to change")
        self.btn_ref_browse.config(text="Change File", bg="#e0f2fe", fg="#0369a1")

        # Auto-detect report type from reference filename
        if detect_report_type:
            try:
                rtype = detect_report_type(self.reference_path)
                self.badge_ref_type.config(
                    text=f"Detected: ✔ {rtype} Report",
                    bg=self.c_brand_light,
                    fg=self.c_brand,
                    font=("Segoe UI", 9, "bold")
                )
                self._append_log(f"[INFO] Reference selected: {basename} (Type: {rtype})\n")
                return
            except Exception:
                pass

        self.badge_ref_type.config(
            text="Detected: ⚠ Unspecified (Auto-detection keyword not in filename)",
            bg="#fef3c7",
            fg="#b45309",
            font=("Segoe UI", 9)
        )
        self._append_log(f"[INFO] Reference selected: {basename}\n")

    def _select_report_type(self, rtype):
        """Toggle segmented pills for Assessment vs Affirmation."""
        self.target_report_type = rtype
        if rtype == "Assessment":
            self.btn_pill_assessment.config(bg=self.c_brand, fg="#ffffff")
            self.btn_pill_affirmation.config(bg="#ffffff", fg=self.c_text_sub)
        else:
            self.btn_pill_affirmation.config(bg=self.c_brand, fg="#ffffff")
            self.btn_pill_assessment.config(bg="#ffffff", fg=self.c_text_sub)

        self._append_log(f"[CONFIG] Target report type set to: {rtype}\n")

    # ========================================================================
    # PIPELINE EXECUTION (BACKGROUND THREAD)
    # ========================================================================

    def _start_pipeline(self):
        if self.is_processing:
            return

        if not self.template_path or not os.path.isfile(self.template_path):
            messagebox.showwarning("Missing Template", "Please select or drop a New Report Template (.docx).")
            return

        if not self.reference_path or not os.path.isfile(self.reference_path):
            messagebox.showwarning("Missing Reference", "Please select or drop a Reference Report (.docx).")
            return

        # Auto-generate draft output name in the template's directory
        tmpl_dir = os.path.dirname(self.template_path)
        tmpl_base, tmpl_ext = os.path.splitext(os.path.basename(self.template_path))
        target_suffix = f"_{self.target_report_type}_Draft"
        output_path = os.path.join(tmpl_dir, f"{tmpl_base}{target_suffix}{tmpl_ext}")

        # Guard: Output file cannot be identical to template
        if os.path.abspath(self.template_path) == os.path.abspath(output_path):
            output_path = os.path.join(tmpl_dir, f"{tmpl_base}_Draft_Output{tmpl_ext}")

        self.is_processing = True
        self.btn_generate.config(state=tk.DISABLED, bg="#94a3b8", cursor="arrow")
        self._disable_footer_buttons()
        self.progress_bar.start(8)
        self.lbl_status.config(text="Processing Word migration in background...")
        self._append_log("\n" + "=" * 60 + f"\n[START] Generating {self.target_report_type} Draft Report...\n" + "=" * 60 + "\n")

        # Launch worker thread
        worker = threading.Thread(
            target=self._run_worker,
            args=(self.template_path, self.reference_path, output_path, self.target_report_type),
            daemon=True
        )
        worker.start()

    def _run_worker(self, template_path, reference_path, output_path, target_type):
        """Worker thread running Word COM automation with stdout redirected."""
        orig_stdout = sys.stdout
        orig_stderr = sys.stderr

        sys.stdout = ThreadSafeLogStream(self.log_queue, orig_stdout)
        sys.stderr = ThreadSafeLogStream(self.log_queue, orig_stderr)

        success = False
        error_msg = None

        try:
            if run_migration is None:
                raise RuntimeError("Migration core engine (phase2_core) could not be loaded.")

            run_migration(
                template_path=template_path,
                reference_path=reference_path,
                output_path=output_path,
                target_report_type=target_type
            )
            success = True
            self.generated_output_path = output_path
        except Exception as exc:
            error_msg = str(exc)
        finally:
            sys.stdout = orig_stdout
            sys.stderr = orig_stderr
            self.after(0, self._on_pipeline_completed, success, error_msg, output_path)

    def _on_pipeline_completed(self, success, error_msg, output_path):
        self.is_processing = False
        self.progress_bar.stop()
        self.btn_generate.config(state=tk.NORMAL, bg=self.c_brand, cursor="hand2")

        if success:
            self.lbl_status.config(text="✔  Draft Report Generated Successfully!")
            self._enable_footer_buttons()

            messagebox.showinfo(
                "Migration Complete",
                f"Your draft {self.target_report_type} report has been successfully generated!\n\n"
                f"Saved to:\n{output_path}\n\n"
                f"You can now open the document in Word to inspect all tracked changes."
            )
        else:
            self.lbl_status.config(text="✖  Process Failed")
            self._append_log(f"\n[ERROR] Pipeline failed: {error_msg}\n")
            messagebox.showerror(
                "Pipeline Error",
                f"An error occurred during report migration:\n\n{error_msg}\n\n"
                f"Please review the activity log for details."
            )

    # ========================================================================
    # LOGGING & QUICK ACTIONS
    # ========================================================================

    def _process_log_queue(self):
        while not self.log_queue.empty():
            try:
                msg = self.log_queue.get_nowait()
                self._append_log(msg)
            except queue.Empty:
                break
        self.after(80, self._process_log_queue)

    def _append_log(self, text):
        self.txt_log.insert(tk.END, text)
        self.txt_log.see(tk.END)

    def _clear_log(self):
        self.txt_log.delete("1.0", tk.END)

    def _enable_footer_buttons(self):
        self.btn_open_word.config(
            bg=self.c_brand,
            fg="#ffffff",
            cursor="hand2"
        )
        self.btn_open_word.bind("<Button-1>", lambda e: self._open_word())

        self.btn_open_folder.config(
            bg=self.c_card,
            fg=self.c_text_main,
            cursor="hand2",
            highlightbackground=self.c_border,
            highlightthickness=1
        )
        self.btn_open_folder.bind("<Button-1>", lambda e: self._open_folder())

    def _disable_footer_buttons(self):
        self.btn_open_word.config(bg="#e2e8f0", fg=self.c_text_muted, cursor="arrow")
        self.btn_open_word.unbind("<Button-1>")
        self.btn_open_folder.config(bg="#e2e8f0", fg=self.c_text_muted, cursor="arrow")
        self.btn_open_folder.unbind("<Button-1>")

    def _open_word(self):
        if self.generated_output_path and os.path.exists(self.generated_output_path):
            try:
                os.startfile(self.generated_output_path)
            except AttributeError:
                # macOS / Linux fallback
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", self.generated_output_path])
            except Exception as exc:
                messagebox.showerror("Error", f"Failed to launch document:\n{exc}")

    def _open_folder(self):
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
# MAIN
# ============================================================================

if __name__ == "__main__":
    app = MRMAutomationApp()
    app.mainloop()
