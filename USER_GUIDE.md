# Model Risk Management (MRM) — Word Report Drafting Tool
### User & Deployment Guide

This tool automates drafting Model Risk Management Word review reports (**Assessment**, **Affirmation**, **Validation**) by intelligently migrating section contents from a finalized Reference Report into a standardized New Template, recording all modifications as **native MS Word Tracked Changes**.

---

## 🚀 Quick Start (For Validators & Colleagues)

### Step 1: Place the Folder on Your Machine
Extract or copy this tool's folder to your bank computer (for example, in your `Documents` or project folder):
```
WordAutomation/
├── gui_app.py
├── phase2_core.py
├── run_tool.bat               <-- Double-click to launch!
└── ...
```

### Step 2: Double-Click `run_tool.bat`
- Automatically checks for Python and the required `pywin32` library.
- Automatically installs `pywin32` if needed.
- Launches the graphical desktop window.

---

## 🖥 How to Use the Tool

1. **Select New Report Template (`.docx`)**
   - Click **Browse...** to select your blank or standardized bank report template.
   - The tool automatically suggests an output draft filename (e.g. `Template_Draft.docx`).

2. **Select Reference Report (`.docx`)**
   - Click **Browse...** to select the prior finalized review report.
   - The tool instantly inspects the filename and displays a green badge:  
     `Detected Type: ✔ Assessment Report` (or `Affirmation` / `Validation`).

3. **Click "▶ Generate Draft Report"**
   - The tool runs MS Word in the background.
   - You can watch the real-time log box as headings are mapped, content is migrated, and red-italic instructions are cleaned up.
   - **Your original template is 100% protected and never modified.**

4. **Review the Results**
   - Click **"📄 Open Draft Document in Word"** to open your new draft.
   - In Word, make sure **Review → All Markup** is selected:
     - **Pasted reference contents** appear as **underlined tracked insertions**.
     - **Default red italic instructions** appear as **struck-through tracked deletions**.
     - **Pre-populated template tables and text** remain **intact and preserved**.

---

## ⚙ Core Automation Logic

| Action | How It Works |
|---|---|
| **Report Detection** | Scans filename for keywords (`Assessment`, `Affirmation`, `Validation`). |
| **Section Matching** | Normalizes heading numbers (e.g., `1.2 Review Scope` → `review scope`) to match headings between template and reference. |
| **Content Migration** | Copies everything between matching headings (paragraphs, tables, images) and inserts it directly under the template heading (bottom-to-top order to prevent shift). |
| **Instruction Cleanup** | Reverse-scans paragraphs and character runs for Red + Italic formatting, deleting them under Track Changes. |
| **File Safety** | Duplicates the template to the output path at the OS level before Word starts. Word never modifies the template. |

---

## 🛠 Troubleshooting

| Issue | Cause | Solution |
|---|---|---|
| *"Cannot write to '...docx'. Please close the document in Word."* | Output file is already open in MS Word Desktop. | Close the Word window and click **Generate Draft Report** again. |
| *"Python was not found on this computer"* | Python is not installed or not in system PATH. | Request Python 3.8+ installation from bank IT. |
| *"pip install pywin32 blocked by proxy"* | Bank firewall restricts internet package downloads. | Ask bank IT to install `pywin32` on your machine (`pip install pywin32 --user`). |
| Ghost `WINWORD.EXE` process | Previous manual Word crash left background process running. | Press `Ctrl + Shift + Esc` (Task Manager), right-click `Microsoft Word` or `WINWORD.EXE`, and select **End Task**. |

