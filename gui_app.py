"""
MRMV Report Content Migration Tool — Desktop GUI
Tailored for Citizens Bank Model Risk Management & Validation (MRMV).

Features:
  - Apple / Microsoft Teams clean minimalist design with Citizens Bank Green (#008450)
  - Trademark-style right-aligned header for "Citizens Model Risk Management & Validation"
  - Side-by-side (左右) upload layout for "New Report Template" and "Reference Report"
  - Native Windows OLE Drag & Drop (ctypes shell32.DragAcceptFiles) + TkinterDnD support
  - Apple-style Segmented Pill Toggle: Assessment vs Affirmation
  - Reference Report auto-detection badge
  - Real Determinate Progress Bar with percentage indicator
  - Custom Apple-style Completion Modal dialog (replaces clunky system alert box)
  - Automated background draft naming and one-click launch in MS Word
"""

import os
import sys
import queue
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext

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

# Optional TkinterDnD import
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    BaseWindow = TkinterDnD.Tk
    HAS_TKDND = True
except Exception:
    BaseWindow = tk.Tk
    HAS_TKDND = False
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
# NATIVE WIN32 CTYPES DRAG & DROP HOOK (ZERO THIRD-PARTY DEPENDENCIES)
# ============================================================================

def setup_native_win32_drag_and_drop(root, on_files_dropped_callback):
    """
    Hook Windows Explorer file drag-and-drop directly via ctypes.
    Ensures dragging files from Windows Explorer shows the '+' (Drop) cursor
    instead of the prohibited 🚫 sign, requiring NO external pip packages.
    """
    if sys.platform != "win32":
        return False

    try:
        import ctypes
        from ctypes import wintypes

        shell32 = ctypes.windll.shell32
        user32 = ctypes.windll.user32

        WM_DROPFILES = 0x0233
        GWLP_WNDPROC = -4

        user32.GetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
        user32.GetWindowLongPtrW.restype = ctypes.c_void_p

        user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
        user32.SetWindowLongPtrW.restype = ctypes.c_void_p

        user32.CallWindowProcW.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
        user32.CallWindowProcW.restype = wintypes.LPARAM

        shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
        shell32.DragAcceptFiles.restype = None

        shell32.DragQueryFileW.argtypes = [ctypes.c_void_p, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]
        shell32.DragQueryFileW.restype = wintypes.UINT

        shell32.DragQueryPoint.argtypes = [ctypes.c_void_p, ctypes.POINTER(wintypes.POINT)]
        shell32.DragQueryPoint.restype = wintypes.BOOL

        shell32.DragFinish.argtypes = [ctypes.c_void_p]
        shell32.DragFinish.restype = None

        WNDPROC = ctypes.WINFUNCTYPE(wintypes.LPARAM, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

        # Get the Win32 HWND for the Tkinter window
        root.update_idletasks()
        try:
            hwnd = int(root.wm_frame(), 16)
        except Exception:
            hwnd = root.winfo_id()

        if not hwnd:
            hwnd = root.winfo_id()

        # Tell Windows Explorer this window accepts dropped files
        shell32.DragAcceptFiles(hwnd, True)

        # Hook window procedure to intercept WM_DROPFILES
        old_wndproc = user32.GetWindowLongPtrW(hwnd, GWLP_WNDPROC)

        def custom_wndproc(h_wnd, msg, w_param, l_param):
            if msg == WM_DROPFILES:
                h_drop = ctypes.c_void_p(w_param)
                count = shell32.DragQueryFileW(h_drop, 0xFFFFFFFF, None, 0)
                file_list = []
                for i in range(count):
                    buf = ctypes.create_unicode_buffer(1024)
                    shell32.DragQueryFileW(h_drop, i, buf, 1024)
                    file_list.append(buf.value)
                pt = wintypes.POINT()
                shell32.DragQueryPoint(h_drop, ctypes.byref(pt))
                shell32.DragFinish(h_drop)

                if file_list:
                    root.after(0, on_files_dropped_callback, file_list, pt.x, pt.y)
                return 0

            return user32.CallWindowProcW(old_wndproc, h_wnd, msg, w_param, l_param)

        root._native_dnd_proc = WNDPROC(custom_wndproc)
        user32.SetWindowLongPtrW(hwnd, GWLP_WNDPROC, ctypes.cast(root._native_dnd_proc, ctypes.c_void_p))
        return True
    except Exception:
        return False


# ============================================================================
# MODERN MODAL DIALOG (APPLE / TEAMS STYLE SUCCESS MODAL)
# ============================================================================

class ModernCompletionModal(tk.Toplevel):
    """Refined Apple/Teams style modal popup dialog for completion notification."""
    def __init__(self, parent, target_type, output_path, on_open_word, on_open_folder):
        super().__init__(parent)
        self.target_type = target_type
        self.output_path = output_path
        self.on_open_word = on_open_word
        self.on_open_folder = on_open_folder

        self.title("Report Generated")
        self.geometry("520x340")
        self.resizable(False, False)
        self.configure(bg="#ffffff")

        # Make modal
        self.transient(parent)
        self.grab_set()

        # Center over parent
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        mx = px + (pw - 520) // 2
        my = py + (ph - 340) // 2
        self.geometry(f"+{max(0, mx)}+{max(0, my)}")

        self._build_content()

        # Key bindings
        self.bind("<Return>", lambda e: self._action_open_word())
        self.bind("<Escape>", lambda e: self.destroy())

    def _build_content(self):
        container = tk.Frame(self, bg="#ffffff", padx=32, pady=28)
        container.pack(fill=tk.BOTH, expand=True)

        # Success Icon Badge
        icon_badge = tk.Label(
            container,
            text="✔",
            font=("Segoe UI", 18, "bold"),
            fg="#008450",
            bg="#e8f5ed",
            width=3,
            height=1
        )
        icon_badge.pack(anchor="w", pady=(0, 10))

        # Title
        title = tk.Label(
            container,
            text="Draft Report Ready",
            font=("Segoe UI", 16, "bold"),
            fg="#0f172a",
            bg="#ffffff"
        )
        title.pack(anchor="w")

        # Subtitle
        subtitle = tk.Label(
            container,
            text=f"Your {self.target_type} report has been successfully drafted with native Track Changes.",
            font=("Segoe UI", 9),
            fg="#475569",
            bg="#ffffff"
        )
        subtitle.pack(anchor="w", pady=(2, 16))

        # File Details Card
        filename = os.path.basename(self.output_path)
        folder = os.path.dirname(self.output_path)

        details_card = tk.Frame(
            container,
            bg="#f8fafc",
            highlightbackground="#e2e8f0",
            highlightthickness=1,
            padx=14,
            pady=10
        )
        details_card.pack(fill=tk.X, pady=(0, 22))

        lbl_file = tk.Label(
            details_card,
            text=f"📄  {filename}",
            font=("Segoe UI", 9, "bold"),
            fg="#0f172a",
            bg="#f8fafc",
            anchor="w"
        )
        lbl_file.pack(fill=tk.X)

        lbl_path = tk.Label(
            details_card,
            text=f"📁  {folder}",
            font=("Segoe UI", 8),
            fg="#64748b",
            bg="#f8fafc",
            anchor="w"
        )
        lbl_path.pack(fill=tk.X, pady=(2, 0))

        # Actions Frame
        btn_frame = tk.Frame(container, bg="#ffffff")
        btn_frame.pack(fill=tk.X)

        btn_word = tk.Button(
            btn_frame,
            text="📄  Open in Word",
            font=("Segoe UI", 10, "bold"),
            bg="#008450",
            fg="#ffffff",
            activebackground="#00663d",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            padx=18,
            pady=8,
            cursor="hand2",
            command=self._action_open_word
        )
        btn_word.pack(side=tk.LEFT, padx=(0, 10))

        btn_folder = tk.Button(
            btn_frame,
            text="📁  View in Folder",
            font=("Segoe UI", 9),
            bg="#f1f5f9",
            fg="#1e293b",
            activebackground="#e2e8f0",
            relief="flat",
            bd=0,
            padx=14,
            pady=8,
            cursor="hand2",
            command=self._action_open_folder
        )
        btn_folder.pack(side=tk.LEFT, padx=(0, 10))

        btn_close = tk.Button(
            btn_frame,
            text="Done",
            font=("Segoe UI", 9),
            bg="#ffffff",
            fg="#64748b",
            relief="flat",
            bd=0,
            padx=12,
            pady=8,
            cursor="hand2",
            command=self.destroy
        )
        btn_close.pack(side=tk.RIGHT)

    def _action_open_word(self):
        self.on_open_word()
        self.destroy()

    def _action_open_folder(self):
        self.on_open_folder()
        self.destroy()


# ============================================================================
# MAIN APPLICATION GUI
# ============================================================================

class MRMAutomationApp(BaseWindow):
    def __init__(self):
        super().__init__()

        self.title("MRMV Report Content Migration Tool")
        self.geometry("960x790")
        self.minsize(880, 700)

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

        # Setup Native Windows Drag & Drop hook (ctypes) + TkinterDnD fallback
        self._init_drag_and_drop()

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
        style.configure(
            "Brand.Horizontal.TProgressbar",
            troughcolor="#e2e8f0",
            background=self.c_brand,
            bordercolor="#e2e8f0",
            lightcolor=self.c_brand,
            darkcolor=self.c_brand
        )

    def _build_ui(self):
        root_padding = tk.Frame(self, bg=self.c_canvas, padx=28, pady=22)
        root_padding.pack(fill=tk.BOTH, expand=True)

        # ---------------- 1. Top Header (Req 1) ----------------
        header_frame = tk.Frame(root_padding, bg=self.c_canvas)
        header_frame.pack(fill=tk.X, pady=(0, 18))

        # Title on Left
        title_label = tk.Label(
            header_frame,
            text="MRMV Report Content Migration Tool",
            font=("Segoe UI", 18, "bold"),
            fg=self.c_text_main,
            bg=self.c_canvas
        )
        title_label.pack(side=tk.LEFT, anchor="w")

        # Subdued Trademark Label on Right (No bullet, regular font, gray)
        brand_label = tk.Label(
            header_frame,
            text="Citizens Model Risk Management & Validation",
            font=("Segoe UI", 9),
            fg="#94a3b8",
            bg=self.c_canvas
        )
        brand_label.pack(side=tk.RIGHT, anchor="e", pady=(6, 0))

        # ---------------- 2. Side-by-Side Upload Cards (左右排版, Req 2) ----------------
        cards_container = tk.Frame(root_padding, bg=self.c_canvas)
        cards_container.pack(fill=tk.X, pady=(0, 16))
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
            text="New Report Template",
            font=("Segoe UI", 11, "bold"),
            fg=self.c_text_main,
            bg=self.c_card
        )
        left_title.pack(side=tk.LEFT)

        # Left Dropzone (Click or Drag & Drop)
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
            text="Reference Report",
            font=("Segoe UI", 11, "bold"),
            fg=self.c_text_main,
            bg=self.c_card
        )
        right_title.pack(side=tk.LEFT)

        # Right Dropzone (Click or Drag & Drop)
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

        self._bind_dropzone_events(self.ref_dropzone, self._browse_reference)

        # Auto-Detection Badge
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

        # ---------------- 3. Central Action & Progress Card (Req 5) ----------------
        action_card = tk.Frame(
            root_padding,
            bg=self.c_card,
            highlightbackground=self.c_border,
            highlightthickness=1,
            padx=18,
            pady=14
        )
        action_card.pack(fill=tk.X, pady=(0, 14))

        action_row = tk.Frame(action_card, bg=self.c_card)
        action_row.pack(fill=tk.X)

        self.btn_generate = tk.Button(
            action_row,
            text="Generate Draft Report",
            font=("Segoe UI", 11, "bold"),
            bg=self.c_brand,
            fg="#ffffff",
            activebackground=self.c_brand_hover,
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            padx=28,
            pady=9,
            cursor="hand2",
            command=self._start_pipeline
        )
        self.btn_generate.pack(side=tk.LEFT)

        # Determinate Progress Bar + Percentage display (Req 5)
        progress_wrapper = tk.Frame(action_row, bg=self.c_card)
        progress_wrapper.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(18, 16))

        progress_info_row = tk.Frame(progress_wrapper, bg=self.c_card)
        progress_info_row.pack(fill=tk.X, pady=(0, 3))

        self.lbl_status = tk.Label(
            progress_info_row,
            text="Ready",
            font=("Segoe UI", 9),
            fg=self.c_text_sub,
            bg=self.c_card
        )
        self.lbl_status.pack(side=tk.LEFT)

        self.lbl_progress_percent = tk.Label(
            progress_info_row,
            text="0%",
            font=("Segoe UI", 9, "bold"),
            fg=self.c_brand,
            bg=self.c_card
        )
        self.lbl_progress_percent.pack(side=tk.RIGHT)

        self.progress_bar = ttk.Progressbar(
            progress_wrapper,
            style="Brand.Horizontal.TProgressbar",
            mode="determinate",
            maximum=100,
            value=0
        )
        self.progress_bar.pack(fill=tk.X)

        # ---------------- 4. Glass-Style Activity Log ----------------
        log_card = tk.Frame(
            root_padding,
            bg="#111827",
            highlightbackground="#1f2937",
            highlightthickness=1,
            padx=14,
            pady=10
        )
        log_card.pack(fill=tk.BOTH, expand=True, pady=(0, 12))

        log_top_bar = tk.Frame(log_card, bg="#111827")
        log_top_bar.pack(fill=tk.X, pady=(0, 4))

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
    # INTERACTION & DRAG-AND-DROP SETUP (Req 3)
    # ========================================================================

    def _bind_dropzone_events(self, dropzone_widget, browse_command):
        """Bind click and hover interactions to a dropzone and all its children."""
        def on_click(e):
            browse_command()

        def on_enter(e):
            dropzone_widget.configure(bg=self.c_drop_hover)
            for child in dropzone_widget.winfo_children():
                if child.cget("text") not in ("Browse Template", "Browse Reference", "Change File"):
                    child.configure(bg=self.c_drop_hover)

        def on_leave(e):
            dropzone_widget.configure(bg=self.c_drop_bg)
            for child in dropzone_widget.winfo_children():
                if child.cget("text") not in ("Browse Template", "Browse Reference", "Change File"):
                    child.configure(bg=self.c_drop_bg)

        dropzone_widget.bind("<Button-1>", on_click)
        dropzone_widget.bind("<Enter>", on_enter)
        dropzone_widget.bind("<Leave>", on_leave)

        for child in dropzone_widget.winfo_children():
            child.bind("<Button-1>", on_click)

    def _init_drag_and_drop(self):
        """
        Initialize drag-and-drop support:
        1. Pure Win32 ctypes hook (DragAcceptFiles) for native Windows Explorer file dragging.
        2. TkinterDnD registration if library is installed.
        """
        # Try native Win32 ctypes hook first (offline, no pip needed)
        success = setup_native_win32_drag_and_drop(self, self._on_native_win32_drop)

        # Also register with TkinterDnD if present
        if HAS_TKDND:
            try:
                self.template_dropzone.drop_target_register(DND_FILES)
                self.template_dropzone.dnd_bind("<<Drop>>", self._on_tkdnd_drop_template)

                self.ref_dropzone.drop_target_register(DND_FILES)
                self.ref_dropzone.dnd_bind("<<Drop>>", self._on_tkdnd_drop_reference)
            except Exception:
                pass

    def _on_native_win32_drop(self, files, drop_x, drop_y):
        """Called when files are dropped from Windows Explorer via ctypes hook."""
        if not files:
            return

        file_path = files[0]
        if not file_path.lower().endswith(".docx"):
            self._append_log("[WARNING] Dropped file is not a .docx document.\n")
            return

        # Determine if dropped on left half (Template) or right half (Reference)
        win_width = self.winfo_width()
        if drop_x < win_width // 2:
            self._set_template_file(file_path)
        else:
            self._set_reference_file(file_path)

    def _clean_drop_path(self, raw_data):
        if not raw_data:
            return None
        path = raw_data.strip()
        if path.startswith("{") and path.endswith("}"):
            path = path[1:-1]
        return path

    def _on_tkdnd_drop_template(self, event):
        path = self._clean_drop_path(event.data)
        if path and path.lower().endswith(".docx") and os.path.isfile(path):
            self._set_template_file(path)

    def _on_tkdnd_drop_reference(self, event):
        path = self._clean_drop_path(event.data)
        if path and path.lower().endswith(".docx") and os.path.isfile(path):
            self._set_reference_file(path)

    # ========================================================================
    # FILE SELECTION & STATE UPDATES (Req 4: removed duplicate "Click to change")
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

        # Update Dropzone appearance (Req 4: only show size, no duplicate "Click to change")
        self.lbl_tmpl_icon.config(text="✔", fg=self.c_brand)
        self.lbl_tmpl_main.config(text=basename[:30] + ("..." if len(basename) > 30 else ""))
        self.lbl_tmpl_sub.config(text=f"{size_kb} KB")
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

        # Update Dropzone appearance (Req 4: only show size, no duplicate "Click to change")
        self.lbl_ref_icon.config(text="✔", fg="#0284c7")
        self.lbl_ref_main.config(text=basename[:30] + ("..." if len(basename) > 30 else ""))
        self.lbl_ref_sub.config(text=f"{size_kb} KB")
        self.btn_ref_browse.config(text="Change File", bg="#e0f2fe", fg="#0369a1")

        # Auto-detect report type
        detected = None
        if detect_report_type:
            try:
                detected = detect_report_type(self.reference_path)
            except Exception:
                pass

        if not detected:
            b_lower = basename.lower()
            for r in ["Assessment", "Affirmation", "Validation"]:
                if r.lower() in b_lower:
                    detected = r
                    break

        if detected:
            self.badge_ref_type.config(
                text=f"Detected: ✔ {detected} Report",
                bg=self.c_brand_light,
                fg=self.c_brand,
                font=("Segoe UI", 9, "bold")
            )
            self._append_log(f"[INFO] Reference selected: {basename} (Type: {detected})\n")
        else:
            self.badge_ref_type.config(
                text="Detected: ⚠ Unspecified (keyword not in filename)",
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
    # PIPELINE EXECUTION & PROGRESS UPDATES (Req 5: Determinate Progress Bar)
    # ========================================================================

    def _update_progress(self, percent, status_text):
        """Thread-safe UI update for determinate progress bar and status label."""
        def apply_update():
            self.progress_bar['value'] = percent
            self.lbl_progress_percent.config(text=f"{percent}%")
            self.lbl_status.config(text=status_text)
        self.after(0, apply_update)

    def _start_pipeline(self):
        if self.is_processing:
            return

        if not self.template_path or not os.path.isfile(self.template_path):
            self._append_log("[WARNING] Please select or drop a New Report Template.\n")
            return

        if not self.reference_path or not os.path.isfile(self.reference_path):
            self._append_log("[WARNING] Please select or drop a Reference Report.\n")
            return

        # Derive clean draft output path automatically in template folder
        tmpl_dir = os.path.dirname(self.template_path)
        tmpl_base, tmpl_ext = os.path.splitext(os.path.basename(self.template_path))
        target_suffix = f"_{self.target_report_type}_Draft"
        output_path = os.path.join(tmpl_dir, f"{tmpl_base}{target_suffix}{tmpl_ext}")

        if os.path.abspath(self.template_path) == os.path.abspath(output_path):
            output_path = os.path.join(tmpl_dir, f"{tmpl_base}_Draft_Output{tmpl_ext}")

        self.is_processing = True
        self.btn_generate.config(state=tk.DISABLED, bg="#94a3b8", cursor="arrow")
        self._disable_footer_buttons()

        # Reset determinate progress bar
        self.progress_bar['value'] = 0
        self.lbl_progress_percent.config(text="0%")
        self.lbl_status.config(text="Starting background migration...")
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
                target_report_type=target_type,
                progress_callback=self._update_progress
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
        """Called on main UI thread once migration worker finishes."""
        self.is_processing = False
        self.btn_generate.config(state=tk.NORMAL, bg=self.c_brand, cursor="hand2")

        if success:
            self.progress_bar['value'] = 100
            self.lbl_progress_percent.config(text="100%")
            self.lbl_status.config(text="✔  Draft Report Generated Successfully!")
            self._enable_footer_buttons()

            # Present modern Apple-style modal dialog (Req 6)
            ModernCompletionModal(
                parent=self,
                target_type=self.target_report_type,
                output_path=output_path,
                on_open_word=self._open_word,
                on_open_folder=self._open_folder
            )
        else:
            self.lbl_status.config(text="✖  Process Failed")
            self._append_log(f"\n[ERROR] Pipeline failed: {error_msg}\n")

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
                self._append_log(f"[ERROR] Failed to launch Word: {exc}\n")

    def _open_folder(self):
        if self.generated_output_path and os.path.exists(self.generated_output_path):
            folder = os.path.dirname(os.path.abspath(self.generated_output_path))
            try:
                if sys.platform == "win32":
                    subprocess.Popen(f'explorer /select,"{os.path.abspath(self.generated_output_path)}"')
                else:
                    subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", folder])
            except Exception as exc:
                self._append_log(f"[ERROR] Failed to open folder: {exc}\n")


# ============================================================================
# MAIN
# ============================================================================

if __name__ == "__main__":
    app = MRMAutomationApp()
    app.mainloop()
