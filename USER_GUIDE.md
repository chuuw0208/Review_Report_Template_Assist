# MRMV Report Content Migration Tool
### User & Deployment Guide — Citizens Bank Model Risk Management & Validation

This desktop application automates drafting Model Risk Management review reports (**Assessment** and **Affirmation**) by migrating section contents from a finalized Reference Report into a standardized New Template, recording all modifications as **native MS Word Tracked Changes**.

---

## 🎨 Design & Experience Highlights

- **Citizens Bank Green Theme**: Branded with Citizens Bank's hallmark green (`#008450`).
- **Instant Branded Startup Screen**: Opens immediately (<0.2s) with a Citizens-branded loading view displaying real-time environment verification (Python runtime, pywin32 COM libraries, migration engine) before smoothly transitioning to the workspace.
- **Proportionally Scaled Down (缩小 1/3)**: Compact `660x370` window geometry that sits comfortably on any desktop without overwhelming the screen.
- **Side-by-Side (左右排版) Layout**:
  - **Left Dropzone**: New Report Template + Target Report Type pill toggle (**Assessment** vs. **Affirmation**).
  - **Right Dropzone**: Reference Report with real-time classification auto-detection badge.
- **Dual Upload Support**: Drag and drop `.docx` files directly from File Explorer, or click anywhere on the dropzone to browse.
- **Compact & Balanced Action Button**: Refined "Generate Draft Report" button with comfortable padding and integrated progress bar.
- **Zero Bottom Clutter**: Redundant bottom grey boxes removed for a clean, distraction-free workflow.
- **Streamlined Completion Modal**: Elegant horizontal checkmark badge with quick access to open the document directly in Word or show it in File Explorer.
- **Intelligent Draft Naming**: Automatically extracts the **Cover Page Title** from the New Report Template and appends `_Draft.docx` (e.g. `Commercial Real Estate PD Model Assessment Report_Draft.docx`).

---

## 🚀 Quick Start (Cross-Platform: Windows & macOS)

### Step 1: Place the Folder on Your Machine
Extract or copy the tool folder to your computer:
```
WordAutomation/
├── Launch_on_Windows.vbs        <-- Double-click on Windows (Instant silent launch, no cmd window!)
├── Launch_on_Mac.command        <-- Double-click on macOS (Launches app directly from Finder!)
├── gui_app.py
├── phase2_core.py
└── USER_GUIDE.md
```

### Step 2: Launch the Tool

- **On Windows (Bank Laptop)**:  
  Double-click **`Launch_on_Windows.vbs`** — opens instantly without any Command Prompt window or flash, shows the startup verification screen, and transitions directly into the workspace.
- **On macOS (MacBook)**:  
  Double-click **`Launch_on_Mac.command`** in Finder — opens the app immediately.


---

## ⚡ Dual-Engine Architecture (Windows COM + Pure OpenXML)

The tool features an intelligent **Dual-Engine Architecture**:
1. **Windows COM Engine**: On Windows laptops with Microsoft Office installed, uses native Word COM automation (`win32com`) to perform tracked changes.
2. **Pure Python OpenXML Engine**: On macOS (MacBook) or systems without MS Word / COM installed, the tool automatically uses its built-in OpenXML engine (`zipfile` + WordprocessingML). It runs in ~0.2 seconds with **zero third-party dependencies**, generating 100% compliant Word documents with native Track Changes (`<w:ins>`, `<w:del>`, `<w:trackRevisions/>`).


---

## 🖥 How to Use the Tool

1. **Upload New Report Template (`.docx`)**
   - **Drag & drop** your blank template file into the left box, or click **Browse Template**.
   - Select your target report type using the segmented pill buttons:
     - `[ Assessment Report ]` — Standard model review report.
     - `[ Affirmation Report ]` — Automatically prunes Section 2 ("Additional Analysis since prior review") and renumbers Section 3 to Section 2 under Track Changes.

2. **Upload Reference Report (`.docx`)**
   - **Drag & drop** your finalized prior review report into the right box, or click **Browse Reference**.
   - The badge automatically detects the report classification:  
     `Detected: ✔ Assessment Report` (or `Affirmation` / `Validation`).

3. **Click "Generate Draft Report"**
   - Word executes safely in the background.
   - The green progress bar smoothly advances from 0% to 100% while status text informs you of each stage.
   - **Your original template is 100% untouched and preserved.**

4. **Review & Open**
   - A success card pops up displaying the detected draft filename and destination.
   - Click **"📄 Open in Word"** to immediately review the draft with Track Changes visible.
   - Click **"📁 View in Folder"** to locate the output file in Windows File Explorer.

---

## ⚙ Core Automation Logic

| Feature | Behavior |
|---|---|
| **Cover Title Naming** | Automatically extracts the title on the cover page of the template, sanitizes it, and appends `_Draft.docx`. |
| **Report Auto-Detection** | Detects reference report type (`Assessment`, `Affirmation`, `Validation`) from the file name. |
| **Section Matching** | Normalizes heading numbers (`1.2 Review Scope` → `review scope`) to match headings between template and reference. |
| **Content Migration** | Migrates paragraphs, tables, and images under each matching heading into the template. |
| **Instruction Cleanup** | Scans for Red + Italic text and deletes it as tracked changes (strikethrough). Normal black text and placeholder tables are preserved. |
| **Affirmation Tailoring** | If Affirmation is selected: <br>1. Section 2 ("Additional Analysis since prior review") and subsections (2.1, 2.2, 2.3) are deleted under Track Changes.<br>2. Section 3 ("3. Governance & Controls", "3.1 Model Controls", "3.2 Performance Monitoring & Reporting", "3.3 Evidence of Contingency Plan") is renumbered to Section 2 ("2.", "2.1", "2.2", "2.3") under Track Changes. |
| **Template Protection** | Creates an isolated working copy before Word starts. Word never modifies the template. |

---

## 🛠 Troubleshooting

| Issue | Cause | Solution |
|---|---|---|
| *"Cannot write to '...docx'. Please close the document in Word."* | Output draft file is currently open in Word. | Close the Word window and click **Generate Draft Report** again. |
| Drag & drop doesn't register | IT environment blocked `tkinterdnd2`. | Built-in Win32 ctypes hook handles native drops, or simply click anywhere on the box to browse. |
| *"Python was not found on this computer"* | Python is not installed or not in PATH. | Request Python 3.8+ from bank IT. |
| Ghost `WINWORD.EXE` process | Previous manual Word crash left background process. | Press `Ctrl + Shift + Esc` (Task Manager), find `Microsoft Word`, and select **End Task**. |
