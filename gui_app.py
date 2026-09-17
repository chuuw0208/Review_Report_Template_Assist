"""
MRMV Report Content Migration Tool — Desktop GUI
Tailored for Citizens Bank Model Risk Management & Validation (MRMV).

Features:
  - Citizens Bank Theme Green (#008450)
  - Instant startup with branded Environment Verification & Loading Screen
  - Scaled-down, compact layout (reduced by 1/3)
  - Side-by-side upload layout for "New Report Template" and "Reference Report"
  - Native Windows OLE Drag & Drop (ctypes shell32.DragAcceptFiles) + TkinterDnD support
  - Segmented Pill Toggle: Assessment vs Affirmation
  - Reference Report auto-detection badge
  - Compact "Generate Draft Report" button + Determinate Progress Bar
  - Streamlined Completion Modal (no Done button, single-row title)
  - Auto-names output document using the Template's cover page title + "_Draft.docx"
"""

import os
import sys
import time
import queue
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

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
# NATIVE WIN32 CTYPES DRAG & DROP HOOK (ZERO THIRD-PARTY DEPENDENCIES)
# ============================================================================

def setup_native_win32_drag_and_drop(root, on_files_dropped_callback):
    """
    Hook Windows Explorer file drag-and-drop directly via ctypes.
    Ensures dragging files from Windows Explorer shows the '+' (Drop) cursor
    and properly delivers dropped file paths without external dependencies.
    Only hooks root.winfo_id() (the client area) so native window title bar
    and window movement are NEVER interfered with.
    """
    if sys.platform != "win32":
        return False

    try:
        import ctypes
        from ctypes import wintypes

        shell32 = ctypes.windll.shell32
        user32 = ctypes.windll.user32

        WM_DROPFILES = 0x0233
        GWL_WNDPROC = -4
        is_64bit = ctypes.sizeof(ctypes.c_void_p) == 8

        root.update_idletasks()
        target_hwnd = root.winfo_id()
        if not target_hwnd:
            return False

        # Configure 64-bit safe ctypes function prototypes
        if is_64bit:
            GetWindowLongPtr = getattr(user32, 'GetWindowLongPtrW', user32.GetWindowLongW)
            SetWindowLongPtr = getattr(user32, 'SetWindowLongPtrW', user32.SetWindowLongW)
        else:
            GetWindowLongPtr = user32.GetWindowLongW
            SetWindowLongPtr = user32.SetWindowLongW

        GetWindowLongPtr.argtypes = [wintypes.HWND, ctypes.c_int]
        GetWindowLongPtr.restype = ctypes.c_void_p

        SetWindowLongPtr.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
        SetWindowLongPtr.restype = ctypes.c_void_p

        user32.CallWindowProcW.argtypes = [
            ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM
        ]
        user32.CallWindowProcW.restype = ctypes.c_longlong if is_64bit else ctypes.c_long

        shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
        shell32.DragQueryFileW.argtypes = [wintypes.HANDLE, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]
        shell32.DragQueryPoint.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.POINT)]
        shell32.DragFinish.argtypes = [wintypes.HANDLE]

        # Enable dragging onto client area ONLY (never touch parent window/title bar!)
        shell32.DragAcceptFiles(target_hwnd, True)

        old_wndproc = GetWindowLongPtr(target_hwnd, GWL_WNDPROC)
        if not old_wndproc:
            return False

        WNDPROC = ctypes.WINFUNCTYPE(
            ctypes.c_longlong if is_64bit else ctypes.c_long,
            wintypes.HWND,
            wintypes.UINT,
            wintypes.WPARAM,
            wintypes.LPARAM
        )

        def custom_wndproc(h_wnd, msg, w_param, l_param):
            if msg == WM_DROPFILES:
                h_drop = w_param
                try:
                    num_files = shell32.DragQueryFileW(h_drop, 0xFFFFFFFF, None, 0)
                    dropped_files = []
                    for i in range(num_files):
                        buf = ctypes.create_unicode_buffer(1024)
                        shell32.DragQueryFileW(h_drop, i, buf, 1024)
                        dropped_files.append(buf.value)

                    pt = wintypes.POINT()
                    shell32.DragQueryPoint(h_drop, ctypes.byref(pt))

                    root.after(0, on_files_dropped_callback, dropped_files, pt.x, pt.y)
                finally:
                    shell32.DragFinish(h_drop)
                return 0

            return user32.CallWindowProcW(old_wndproc, h_wnd, msg, w_param, l_param)

        new_proc = WNDPROC(custom_wndproc)
        res = SetWindowLongPtr(target_hwnd, GWL_WNDPROC, ctypes.cast(new_proc, ctypes.c_void_p))
        if res:
            root._native_dnd_proc = new_proc
            return True
        return False

    except Exception as ex:
        print(f"[DragAndDrop] Native hook failed: {ex}")
        return False


# ============================================================================
# MODERN COMPLETION MODAL DIALOG (COMPACT & CLEAN)
# ============================================================================

class ModernCompletionModal(tk.Toplevel):
    """Refined Apple/Teams style modal dialog for completion notification."""
    def __init__(self, parent, target_type, output_path, on_open_word, on_open_folder):
        super().__init__(parent)
        self.target_type = target_type
        self.output_path = output_path
        self.on_open_word = on_open_word
        self.on_open_folder = on_open_folder

        self.title("Draft Generated")
        self.geometry("460x220")
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
        mx = px + (pw - 460) // 2
        my = py + (ph - 220) // 2
        self.geometry(f"+{max(0, mx)}+{max(0, my)}")

        self._build_content()

        # Keyboard shortcuts
        self.bind("<Return>", lambda e: self._action_open_word())
        self.bind("<Escape>", lambda e: self.destroy())

    def _build_content(self):
        container = tk.Frame(self, bg="#ffffff", padx=24, pady=18)
        container.pack(fill=tk.BOTH, expand=True)

        # Header Row: Checkmark badge on left, Title on right
        header_row = tk.Frame(container, bg="#ffffff")
        header_row.pack(fill=tk.X, pady=(0, 6))

        icon_badge = tk.Label(
            header_row,
            text="✔",
            font=("Segoe UI", 11, "bold"),
            fg="#008450",
            bg="#e8f5ed",
            width=2,
            height=1
        )
        icon_badge.pack(side=tk.LEFT, padx=(0, 8))

        title = tk.Label(
            header_row,
            text="Draft Report Ready",
            font=("Segoe UI", 13, "bold"),
            fg="#0f172a",
            bg="#ffffff"
        )
        title.pack(side=tk.LEFT)

        # Allow dragging modal dialog by clicking its header
        def _on_modal_drag_start(e):
            self._modal_drag_x = e.x_root - self.winfo_x()
            self._modal_drag_y = e.y_root - self.winfo_y()

        def _on_modal_drag_motion(e):
            nx = e.x_root - self._modal_drag_x
            ny = e.y_root - self._modal_drag_y
            self.geometry(f"+{nx}+{ny}")

        header_row.bind("<Button-1>", _on_modal_drag_start)
        header_row.bind("<B1-Motion>", _on_modal_drag_motion)
        title.bind("<Button-1>", _on_modal_drag_start)
        title.bind("<B1-Motion>", _on_modal_drag_motion)
        icon_badge.bind("<Button-1>", _on_modal_drag_start)
        icon_badge.bind("<B1-Motion>", _on_modal_drag_motion)

        # Subtitle
        subtitle = tk.Label(
            container,
            text=f"Your {self.target_type} draft has been generated with native Track Changes.",
            font=("Segoe UI", 8),
            fg="#475569",
            bg="#ffffff"
        )
        subtitle.pack(anchor="w", pady=(0, 12))

        # File Details Box
        filename = os.path.basename(self.output_path)
        folder = os.path.dirname(self.output_path)

        details_card = tk.Frame(
            container,
            bg="#f8fafc",
            highlightbackground="#e2e8f0",
            highlightthickness=1,
            padx=10,
            pady=8
        )
        details_card.pack(fill=tk.X, pady=(0, 14))

        lbl_file = tk.Label(
            details_card,
            text=f"📄  {filename}",
            font=("Segoe UI", 8, "bold"),
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

        # Actions Row
        btn_frame = tk.Frame(container, bg="#ffffff")
        btn_frame.pack(fill=tk.X)

        btn_word = tk.Button(
            btn_frame,
            text="📄  Open in Word",
            font=("Segoe UI", 9, "bold"),
            bg="#008450",
            fg="#ffffff",
            activebackground="#00663d",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            padx=14,
            pady=6,
            cursor="hand2",
            command=self._action_open_word
        )
        btn_word.pack(side=tk.LEFT, padx=(0, 8))

        btn_folder = tk.Button(
            btn_frame,
            text="📁  View in Folder",
            font=("Segoe UI", 8),
            bg="#f1f5f9",
            fg="#1e293b",
            activebackground="#e2e8f0",
            relief="flat",
            bd=0,
            padx=12,
            pady=6,
            cursor="hand2",
            command=self._action_open_folder
        )
        btn_folder.pack(side=tk.LEFT)

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
        # Scaled down by ~1/3 (from 940x530 to 660x370)
        self.geometry("660x370")
        self.minsize(640, 350)

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
        self.is_processing = False
        self.template_path = None
        self.reference_path = None
        self.generated_output_path = None
        self.target_report_type = "Assessment"  # Default: Assessment

        self.startup_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        self._configure_styles()

        # Display immediate branded Startup / Loading screen while checking environment
        self._build_startup_ui()
        self.after(30, self._process_startup_queue)
        threading.Thread(target=self._run_startup_checks, daemon=True).start()

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

    # ========================================================================
    # 0. STARTUP / VERIFICATION VIEW (IMMEDIATE FEEDBACK)
    # ========================================================================

    def _build_startup_ui(self):
        """Display an immediate, elegant splash/loading screen while checking environment."""
        self.startup_frame = tk.Frame(self, bg=self.c_canvas)
        self.startup_frame.pack(fill=tk.BOTH, expand=True)

        card = tk.Frame(
            self.startup_frame,
            bg=self.c_card,
            highlightbackground=self.c_border,
            highlightthickness=1,
            padx=28,
            pady=22
        )
        card.place(relx=0.5, rely=0.5, anchor="center")

        # Top Badge
        badge_frame = tk.Frame(card, bg=self.c_card)
        badge_frame.pack(anchor="w", pady=(0, 6))

        badge = tk.Label(
            badge_frame,
            text="CITIZENS BANK · MRMV",
            font=("Segoe UI", 7, "bold"),
            fg=self.c_brand,
            bg=self.c_brand_light,
            padx=7,
            pady=2
        )
        badge.pack(side=tk.LEFT)

        # Title & Subtitle
        title = tk.Label(
            card,
            text="MRMV Report Content Migration Tool",
            font=("Segoe UI", 12, "bold"),
            fg=self.c_brand,
            bg=self.c_card
        )
        title.pack(anchor="w")

        subtitle = tk.Label(
            card,
            text="Initializing automation environment & requirements...",
            font=("Segoe UI", 8),
            fg=self.c_text_sub,
            bg=self.c_card
        )
        subtitle.pack(anchor="w", pady=(2, 14))

        # Allow dragging the window during startup view
        def _on_startup_drag_start(e):
            self._startup_drag_x = e.x_root - self.winfo_x()
            self._startup_drag_y = e.y_root - self.winfo_y()

        def _on_startup_drag_motion(e):
            nx = e.x_root - self._startup_drag_x
            ny = e.y_root - self._startup_drag_y
            self.geometry(f"+{nx}+{ny}")

        self.startup_frame.bind("<Button-1>", _on_startup_drag_start)
        self.startup_frame.bind("<B1-Motion>", _on_startup_drag_motion)
        card.bind("<Button-1>", _on_startup_drag_start)
        card.bind("<B1-Motion>", _on_startup_drag_motion)
        title.bind("<Button-1>", _on_startup_drag_start)
        title.bind("<B1-Motion>", _on_startup_drag_motion)
        subtitle.bind("<Button-1>", _on_startup_drag_start)
        subtitle.bind("<B1-Motion>", _on_startup_drag_motion)

        # Progress Bar
        self.startup_progress = ttk.Progressbar(
            card,
            style="Brand.Horizontal.TProgressbar",
            mode="determinate",
            maximum=100,
            value=15,
            length=320
        )
        self.startup_progress.pack(fill=tk.X, pady=(0, 10))

        # Status Checklist
        self.lbl_step_python = tk.Label(
            card,
            text="⏳  Checking Python runtime environment...",
            font=("Segoe UI", 8),
            fg=self.c_text_sub,
            bg=self.c_card,
            anchor="w"
        )
        self.lbl_step_python.pack(fill=tk.X, pady=1)

        self.lbl_step_win32 = tk.Label(
            card,
            text="⏳  Checking pywin32 & COM libraries...",
            font=("Segoe UI", 8),
            fg=self.c_text_sub,
            bg=self.c_card,
            anchor="w"
        )
        self.lbl_step_win32.pack(fill=tk.X, pady=1)

        self.lbl_step_engine = tk.Label(
            card,
            text="⏳  Loading document migration engine...",
            font=("Segoe UI", 8),
            fg=self.c_text_sub,
            bg=self.c_card,
            anchor="w"
        )
        self.lbl_step_engine.pack(fill=tk.X, pady=1)

    def _process_startup_queue(self):
        """Thread-safe queue reader on the main Tkinter thread."""
        try:
            while True:
                item = self.startup_queue.get_nowait()
                step, text, fg, progress_val = item
                if step == "done":
                    self._transition_to_main_ui()
                    return
                elif step == "python" and hasattr(self, 'lbl_step_python'):
                    self.lbl_step_python.config(text=text, fg=fg)
                elif step == "win32" and hasattr(self, 'lbl_step_win32'):
                    self.lbl_step_win32.config(text=text, fg=fg)
                elif step == "engine" and hasattr(self, 'lbl_step_engine'):
                    self.lbl_step_engine.config(text=text, fg=fg)
                if progress_val is not None and hasattr(self, 'startup_progress'):
                    self.startup_progress['value'] = progress_val
        except queue.Empty:
            pass
        if hasattr(self, 'startup_frame') and self.startup_frame:
            self.after(30, self._process_startup_queue)

    def _run_startup_checks(self):
        """Perform environment verification in the background and update startup UI."""
        # Step 1: Python environment
        time.sleep(0.15)
        py_ver = f"Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
        self.startup_queue.put(("python", f"✔  Python Runtime: {py_ver}", self.c_brand, 40))

        # Step 2: Automation Interface (Windows COM vs macOS)
        time.sleep(0.2)
        if sys.platform == "win32":
            has_win32 = False
            try:
                import win32com.client
                import pythoncom
                has_win32 = True
            except ImportError:
                try:
                    self.startup_queue.put(("win32", "⚙  Installing 'pywin32' dependency...", "#d97706", None))
                    subprocess.run(
                        [sys.executable, "-m", "pip", "install", "pywin32", "--quiet"],
                        capture_output=True,
                        timeout=30
                    )
                    import win32com.client
                    import pythoncom
                    has_win32 = True
                except Exception:
                    has_win32 = False

            if has_win32:
                self.startup_queue.put(("win32", "✔  Word COM Interface: Ready (pywin32)", self.c_brand, 75))
            else:
                self.startup_queue.put(("win32", "⚠  Word COM: pywin32 not loaded (migration will warn)", "#d97706", 75))
        else:
            # macOS platform
            self.startup_queue.put(("win32", "✔  macOS Platform Detected: Ready", self.c_brand, 75))

        # Step 3: Migration engine
        time.sleep(0.2)
        global run_migration, detect_report_type
        if run_migration is None:
            try:
                from phase2_core import run_migration as rm, detect_report_type as drt
                run_migration = rm
                detect_report_type = drt
            except Exception:
                pass

        if run_migration is not None:
            self.startup_queue.put(("engine", "✔  Document Migration Engine: Ready", self.c_brand, 100))
        else:
            self.startup_queue.put(("engine", "⚠  Document Migration Engine: Failed to import", "#dc2626", 100))

        # Short pause so user sees all items checked, then transition to upload screen
        time.sleep(0.35)
        self.startup_queue.put(("done", None, None, None))

    def _transition_to_main_ui(self):
        """Smoothly switch from startup checklist to the main upload workspace."""
        if hasattr(self, 'startup_frame') and self.startup_frame:
            self.startup_frame.destroy()
            self.startup_frame = None

        self._build_ui()
        self._init_drag_and_drop()

    # ========================================================================
    # 1. MAIN UPLOAD WORKSPACE (SCALED DOWN BY ~1/3, NO FOOTER BOXES)
    # ========================================================================

    def _build_ui(self):
        root_padding = tk.Frame(self, bg=self.c_canvas, padx=16, pady=12)
        root_padding.pack(fill=tk.BOTH, expand=True)

        # ---------------- 1. Top Header ----------------
        header_frame = tk.Frame(root_padding, bg=self.c_canvas)
        header_frame.pack(fill=tk.X, pady=(0, 10))

        # Title in Theme Green (#008450)
        title_label = tk.Label(
            header_frame,
            text="MRMV Report Content Migration Tool",
            font=("Segoe UI", 13, "bold"),
            fg=self.c_brand,
            bg=self.c_canvas
        )
        title_label.pack(side=tk.LEFT, anchor="w")

        # Allow dragging the window by clicking the top header inside the UI
        def _on_main_drag_start(e):
            self._main_drag_x = e.x_root - self.winfo_x()
            self._main_drag_y = e.y_root - self.winfo_y()

        def _on_main_drag_motion(e):
            nx = e.x_root - self._main_drag_x
            ny = e.y_root - self._main_drag_y
            self.geometry(f"+{nx}+{ny}")

        header_frame.bind("<Button-1>", _on_main_drag_start)
        header_frame.bind("<B1-Motion>", _on_main_drag_motion)
        title_label.bind("<Button-1>", _on_main_drag_start)
        title_label.bind("<B1-Motion>", _on_main_drag_motion)

        # ---------------- 2. Side-by-Side Upload Cards (左右排版) ----------------
        cards_container = tk.Frame(root_padding, bg=self.c_canvas)
        cards_container.pack(fill=tk.X, pady=(0, 10))
        cards_container.columnconfigure(0, weight=1, uniform="group1")
        cards_container.columnconfigure(1, weight=1, uniform="group1")

        # --- LEFT CARD: New Report Template ---
        self.left_card = tk.Frame(
            cards_container,
            bg=self.c_card,
            highlightbackground=self.c_border,
            highlightthickness=1,
            padx=12,
            pady=10
        )
        self.left_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        left_header = tk.Frame(self.left_card, bg=self.c_card)
        left_header.pack(fill=tk.X, pady=(0, 6))

        left_title = tk.Label(
            left_header,
            text="New Report Template",
            font=("Segoe UI", 10, "bold"),
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
            padx=10,
            pady=10,
            cursor="hand2"
        )
        self.template_dropzone.pack(fill=tk.X, pady=(0, 8))

        self.lbl_tmpl_icon = tk.Label(
            self.template_dropzone,
            text="📄",
            font=("Segoe UI Emoji", 20),
            bg=self.c_drop_bg,
            fg=self.c_brand
        )
        self.lbl_tmpl_icon.pack(pady=(0, 2))

        self.lbl_tmpl_main = tk.Label(
            self.template_dropzone,
            text="Drag & drop template here",
            font=("Segoe UI", 9, "bold"),
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
        self.lbl_tmpl_sub.pack(pady=(1, 4))

        self.btn_tmpl_browse = tk.Label(
            self.template_dropzone,
            text="Browse Template",
            font=("Segoe UI", 8, "bold"),
            fg=self.c_brand,
            bg=self.c_brand_light,
            padx=10,
            pady=3,
            cursor="hand2"
        )
        self.btn_tmpl_browse.pack()

        self._bind_dropzone_events(self.template_dropzone, self._browse_template)

        # Target Report Type Toggle Section
        target_section = tk.Frame(self.left_card, bg=self.c_card)
        target_section.pack(fill=tk.X, pady=(2, 0))

        target_label = tk.Label(
            target_section,
            text="Desired Output Report Type:",
            font=("Segoe UI", 8, "bold"),
            fg=self.c_text_main,
            bg=self.c_card
        )
        target_label.pack(anchor="w", pady=(0, 4))

        # Apple-style Segmented Pill Control
        segmented_frame = tk.Frame(target_section, bg="#e2e8f0", padx=2, pady=2)
        segmented_frame.pack(fill=tk.X)

        self.btn_pill_assessment = tk.Label(
            segmented_frame,
            text="Assessment Report",
            font=("Segoe UI", 8, "bold"),
            bg=self.c_brand,
            fg="#ffffff",
            padx=10,
            pady=4,
            cursor="hand2"
        )
        self.btn_pill_assessment.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))

        self.btn_pill_affirmation = tk.Label(
            segmented_frame,
            text="Affirmation Report",
            font=("Segoe UI", 8, "bold"),
            bg="#ffffff",
            fg=self.c_text_sub,
            padx=10,
            pady=4,
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
            padx=12,
            pady=10
        )
        self.right_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        right_header = tk.Frame(self.right_card, bg=self.c_card)
        right_header.pack(fill=tk.X, pady=(0, 6))

        right_title = tk.Label(
            right_header,
            text="Reference Report",
            font=("Segoe UI", 10, "bold"),
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
            padx=10,
            pady=10,
            cursor="hand2"
        )
        self.ref_dropzone.pack(fill=tk.X, pady=(0, 8))

        self.lbl_ref_icon = tk.Label(
            self.ref_dropzone,
            text="📑",
            font=("Segoe UI Emoji", 20),
            bg=self.c_drop_bg,
            fg="#0284c7"
        )
        self.lbl_ref_icon.pack(pady=(0, 2))

        self.lbl_ref_main = tk.Label(
            self.ref_dropzone,
            text="Drag & drop reference here",
            font=("Segoe UI", 9, "bold"),
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
        self.lbl_ref_sub.pack(pady=(1, 4))

        self.btn_ref_browse = tk.Label(
            self.ref_dropzone,
            text="Browse Reference",
            font=("Segoe UI", 8, "bold"),
            fg="#0369a1",
            bg="#e0f2fe",
            padx=10,
            pady=3,
            cursor="hand2"
        )
        self.btn_ref_browse.pack()

        self._bind_dropzone_events(self.ref_dropzone, self._browse_reference)

        # Auto-Detection Badge
        ref_status_section = tk.Frame(self.right_card, bg=self.c_card)
        ref_status_section.pack(fill=tk.X, pady=(2, 0))

        ref_status_title = tk.Label(
            ref_status_section,
            text="Reference Classification:",
            font=("Segoe UI", 8, "bold"),
            fg=self.c_text_main,
            bg=self.c_card
        )
        ref_status_title.pack(anchor="w", pady=(0, 4))

        self.badge_ref_type = tk.Label(
            ref_status_section,
            text="Type: Auto-detected from filename",
            font=("Segoe UI", 8),
            bg="#f1f5f9",
            fg=self.c_text_sub,
            padx=8,
            pady=5
        )
        self.badge_ref_type.pack(fill=tk.X)

        # ---------------- 3. Central Action & Progress Card ----------------
        action_card = tk.Frame(
            root_padding,
            bg=self.c_card,
            highlightbackground=self.c_border,
            highlightthickness=1,
            padx=12,
            pady=10
        )
        action_card.pack(fill=tk.X)

        action_row = tk.Frame(action_card, bg=self.c_card)
        action_row.pack(fill=tk.X)

        # Compact Generate Draft Report button (scaled down, comfortable padding)
        self.btn_generate = tk.Button(
            action_row,
            text="Generate Draft Report",
            font=("Segoe UI", 10, "bold"),
            bg=self.c_brand,
            fg="#ffffff",
            activebackground=self.c_brand_hover,
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            padx=20,
            pady=8,
            cursor="hand2",
            command=self._start_pipeline
        )
        self.btn_generate.pack(side=tk.LEFT)

        # Determinate Progress Bar + Percentage display
        progress_wrapper = tk.Frame(action_row, bg=self.c_card)
        progress_wrapper.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(16, 8), pady=2)

        progress_info_row = tk.Frame(progress_wrapper, bg=self.c_card)
        progress_info_row.pack(fill=tk.X, pady=(0, 3))

        self.lbl_status = tk.Label(
            progress_info_row,
            text="Ready",
            font=("Segoe UI", 8),
            fg=self.c_text_sub,
            bg=self.c_card
        )
        self.lbl_status.pack(side=tk.LEFT)

        self.lbl_progress_percent = tk.Label(
            progress_info_row,
            text="0%",
            font=("Segoe UI", 8, "bold"),
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
        setup_native_win32_drag_and_drop(self, self._on_native_win32_drop)

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

        self.lbl_tmpl_icon.config(text="✔", fg=self.c_brand)
        self.lbl_tmpl_main.config(text=basename[:24] + ("..." if len(basename) > 24 else ""))
        self.lbl_tmpl_sub.config(text=f"{size_kb} KB")
        self.btn_tmpl_browse.config(text="Change File", bg="#d1fae5", fg=self.c_brand)

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
        self.lbl_ref_main.config(text=basename[:24] + ("..." if len(basename) > 24 else ""))
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
                font=("Segoe UI", 8, "bold")
            )
        else:
            self.badge_ref_type.config(
                text="Detected: ⚠ Unspecified",
                bg="#fef3c7",
                fg="#b45309",
                font=("Segoe UI", 8)
            )

    def _select_report_type(self, rtype):
        """Toggle segmented pills for Assessment vs Affirmation."""
        self.target_report_type = rtype
        if rtype == "Assessment":
            self.btn_pill_assessment.config(bg=self.c_brand, fg="#ffffff")
            self.btn_pill_affirmation.config(bg="#ffffff", fg=self.c_text_sub)
        else:
            self.btn_pill_affirmation.config(bg=self.c_brand, fg="#ffffff")
            self.btn_pill_assessment.config(bg="#ffffff", fg=self.c_text_sub)

    # ========================================================================
    # PIPELINE EXECUTION & PROGRESS UPDATES
    # ========================================================================

    def _update_progress(self, percent, status_text):
        """Thread-safe worker callback pushing updates into progress_queue."""
        self.progress_queue.put(("progress", percent, status_text))

    def _process_pipeline_queue(self):
        """Thread-safe queue reader for migration progress on the main Tkinter thread."""
        try:
            while True:
                item = self.progress_queue.get_nowait()
                tag = item[0]
                if tag == "progress":
                    _, percent, status_text = item
                    self.progress_bar['value'] = percent
                    self.lbl_progress_percent.config(text=f"{percent}%")
                    self.lbl_status.config(text=status_text)
                elif tag == "done":
                    _, success, error_msg, result_output_path = item
                    self._on_pipeline_completed(success, error_msg, result_output_path)
                    return
        except queue.Empty:
            pass
        if self.is_processing:
            self.after(30, self._process_pipeline_queue)

    def _start_pipeline(self):
        """Validates inputs and spawns migration worker in background thread."""
        if self.is_processing:
            return

        if not self.template_path or not os.path.isfile(self.template_path):
            messagebox.showwarning("Missing Template", "Please select or drop a New Report Template (.docx).")
            return

        if not self.reference_path or not os.path.isfile(self.reference_path):
            messagebox.showwarning("Missing Reference", "Please select or drop a Reference Report (.docx).")
            return

        self.is_processing = True
        self.btn_generate.config(state=tk.DISABLED, bg="#94a3b8", cursor="arrow")

        # Reset determinate progress bar
        self.progress_bar['value'] = 0
        self.lbl_progress_percent.config(text="0%")
        self.lbl_status.config(text="Starting migration pipeline...")
        self.after(30, self._process_pipeline_queue)

        # Launch worker thread (output_path=None tells run_migration to extract cover title)
        worker = threading.Thread(
            target=self._run_worker,
            args=(self.template_path, self.reference_path, self.target_report_type),
            daemon=True
        )
        worker.start()

    def _run_worker(self, template_path, reference_path, target_type):
        """Worker thread executing Word COM automation in the background."""
        success = False
        error_msg = None
        result_output_path = None

        try:
            if run_migration is None:
                raise RuntimeError("Migration core engine (phase2_core) could not be loaded.")

            result_output_path = run_migration(
                template_path=template_path,
                reference_path=reference_path,
                output_path=None,
                target_report_type=target_type,
                progress_callback=self._update_progress
            )
            success = True
            self.generated_output_path = result_output_path
        except Exception as exc:
            error_msg = str(exc)
        finally:
            self.progress_queue.put(("done", success, error_msg, result_output_path))

    def _on_pipeline_completed(self, success, error_msg, output_path):
        """Called on main UI thread once migration worker finishes."""
        self.is_processing = False
        self.btn_generate.config(state=tk.NORMAL, bg=self.c_brand, cursor="hand2")

        if success:
            self.progress_bar['value'] = 100
            self.lbl_progress_percent.config(text="100%")
            self.lbl_status.config(text="✔  Draft Report Generated Successfully!")

            # Present modern Apple-style modal dialog
            ModernCompletionModal(
                parent=self,
                target_type=self.target_report_type,
                output_path=output_path,
                on_open_word=self._open_word,
                on_open_folder=self._open_folder
            )
        else:
            self.lbl_status.config(text="✖  Process Failed")
            messagebox.showerror("Migration Error", f"An error occurred during report migration:\n\n{error_msg}")

    # ========================================================================
    # QUICK ACTIONS (CALLED BY COMPLETION MODAL)
    # ========================================================================

    def _open_word(self):
        if self.generated_output_path and os.path.exists(self.generated_output_path):
            try:
                os.startfile(self.generated_output_path)
            except AttributeError:
                subprocess.Popen(["open", self.generated_output_path])

    def _open_folder(self):
        if self.generated_output_path and os.path.exists(self.generated_output_path):
            folder = os.path.dirname(self.generated_output_path)
            try:
                os.startfile(folder)
            except AttributeError:
                subprocess.Popen(["open", folder])


# ============================================================================
# ENTRY POINT
# ============================================================================

def main():
    app = MRMAutomationApp()
    app.mainloop()


if __name__ == "__main__":
    main()
