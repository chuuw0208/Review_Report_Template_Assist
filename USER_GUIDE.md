# MRMV Report Content Migration Tool
### User & Deployment Guide — Citizens Bank Model Risk Management & Validation

This desktop application automates drafting Model Risk Management review reports (**Assessment** and **Affirmation**) by migrating section contents from a finalized Reference Report into a standardized New Template, recording all modifications as **native MS Word Tracked Changes**.

---

## 🎨 Design & Experience Highlights

- **Citizens Bank Green Theme**: Branded with Citizens Bank's hallmark green (`#008450`).
- **Clean Minimalist Architecture**: Pure Apple/Teams card layout without terminal/log clutter.
- **Side-by-Side (左右排版) Layout**:
  - **Left Dropzone**: New Report Template + Target Report Type pill toggle (**Assessment** vs. **Affirmation**).
  - **Right Dropzone**: Reference Report with real-time classification auto-detection badge.
- **Dual Upload Support**: Drag and drop `.docx` files directly from File Explorer, or click anywhere on the dropzone to browse.
- **Intelligent Draft Naming**: Automatically extracts the **Cover Page Title** from the New Report Template and appends `_Draft.docx` (e.g. `Commercial Real Estate PD Model Assessment Report_Draft.docx`).

---

## 🚀 Quick Start (For Validators & Colleagues)

### Step 1: Place the Folder on Your Machine
Extract or copy the tool folder to your computer:
```
WordAutomation/
├── gui_app.py
├── phase2_core.py
├── run_tool.bat               <-- Double-click to launch!
└── USER_GUIDE.md
```

### Step 2: Double-Click `run_tool.bat`
- Verifies your Python installation and `pywin32`.
- Launches the modern desktop application.

---

## 🖥 How to Use the Tool

1. **Upload New Report Template (`.docx`)**
   - **Drag & drop** your blank template file into the left box, or click **Browse Template**.
   - Select your target report type using the segmented pill buttons:
     - `[ Assessment Report ]` — Standard model review report.
     - `[ Affirmation Report ]` — Automatically prunes Section 2 ("Additional Analysis since prior review") and its subsections under Track Changes.

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
| **Affirmation Tailoring** | If Affirmation is selected, Section 2 ("Additional Analysis since prior review") and its subsections (2.1, 2.2, 2.3) are deleted under Track Changes. |
| **Template Protection** | Creates an isolated working copy before Word starts. Word never modifies the template. |

---

## 🛠 Troubleshooting

| Issue | Cause | Solution |
|---|---|---|
| *"Cannot write to '...docx'. Please close the document in Word."* | Output draft file is currently open in Word. | Close the Word window and click **Generate Draft Report** again. |
| Drag & drop doesn't register | IT environment blocked `tkinterdnd2`. | Built-in Win32 ctypes hook handles native drops, or simply click anywhere on the box to browse. |
| *"Python was not found on this computer"* | Python is not installed or not in PATH. | Request Python 3.8+ from bank IT. |
| Ghost `WINWORD.EXE` process | Previous manual Word crash left background process. | Press `Ctrl + Shift + Esc` (Task Manager), find `Microsoft Word`, and select **End Task**. |
