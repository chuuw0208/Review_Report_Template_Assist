# MRMV Report Content Migration Tool
### User & Deployment Guide — Citizens Bank Model Risk Management & Validation

This desktop application automates drafting Model Risk Management review reports (**Assessment** and **Affirmation**) by migrating section contents from a finalized Reference Report into a standardized New Template, recording all modifications as **native MS Word Tracked Changes**.

---

## 🎨 Design & Experience Highlights

- **Citizens Bank Green Theme**: Branded with Citizens Bank's hallmark green (`#008450`) and soft mint accents (`#e8f5ed`).
- **Apple / Microsoft Teams Modern Aesthetic**: Minimalist, clean card containers with generous spacing, refined typography, and glass-feel activity log.
- **Side-by-Side (左右排版) Layout**:
  - **Left Dropzone**: New Report Template + Target Report Type pill toggle (**Assessment** vs. **Affirmation**).
  - **Right Dropzone**: Reference Report with real-time classification auto-detection.
- **Dual Upload Support**: Drag and drop `.docx` files directly from File Explorer, or click anywhere on the dropzone to browse.
- **Automated Output Management**: Automatically names and saves the draft report in the template's folder (e.g. `Template_Assessment_Draft.docx` or `Template_Affirmation_Draft.docx`).

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
- Verifies your Python installation.
- Automatically verifies and sets up `pywin32` (and optional `tkinterdnd2` for drag-and-drop).
- Launches the modern desktop application.

---

## 🖥 How to Use the Tool

1. **Upload New Report Template (`.docx`)**
   - **Drag & drop** your blank template file into the left box, or click **Browse Template**.
   - Select your target report type using the segmented pill buttons:
     - `[ Assessment Report ]` — Keeps all standard sections including Section 2 ("Additional Analysis since prior review").
     - `[ Affirmation Report ]` — Automatically prunes Section 2 and its subsections under Track Changes.

2. **Upload Reference Report (`.docx`)**
   - **Drag & drop** your finalized prior review report into the right box, or click **Browse Reference**.
   - The badge instantly detects the report classification:  
     `Detected: ✔ Assessment Report` (or `Affirmation` / `Validation`).

3. **Click "Generate Draft Report"**
   - Word runs securely in the background.
   - The activity console displays live updates as sections are mapped, text/tables are migrated, and red-italic instructions are cleaned up.
   - **Your original template is 100% untouched and preserved.**

4. **Review & Open**
   - Click **"📄 Open Draft Document in Word"** to review the draft with Track Changes enabled.
   - Click **"📁 Open Containing Folder"** to jump straight to the output file in Windows File Explorer.

---

## ⚙ Core Automation Logic

| Feature | Behavior |
|---|---|
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
| Drag & drop doesn't register | IT environment blocked `tkinterdnd2`. | Simply click anywhere on the upload box to browse and select the file. |
| *"Python was not found on this computer"* | Python is not installed or not in PATH. | Request Python 3.8+ from bank IT. |
| Ghost `WINWORD.EXE` process | Previous manual Word crash left background process. | Press `Ctrl + Shift + Esc` (Task Manager), find `Microsoft Word`, and select **End Task**. |
