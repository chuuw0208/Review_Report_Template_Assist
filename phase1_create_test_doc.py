"""
Phase 1 – Helper: Create a test Word document for the PoC.

Run this FIRST to generate 'test_input.docx' in the same directory.
It uses COM so it also serves as a quick smoke-test that pywin32 works.

Usage:
    python phase1_create_test_doc.py
"""

import os
import sys
import win32com.client
import pythoncom

def create_test_document():
    """Create a minimal Word doc with placeholder text for the PoC."""
    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               "test_input.docx")

    word = None
    doc = None
    try:
        # Initialize COM for this thread
        pythoncom.CoInitialize()

        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False          # run invisibly
        word.DisplayAlerts = False    # suppress all dialogs

        doc = word.Documents.Add()    # blank document

        # Insert some sample content
        rng = doc.Content
        rng.Text = (
            "This is the first paragraph of the test document.\r\n"
            "\r\n"
            "Delete Me\r\n"
            "\r\n"
            "This paragraph should remain after the PoC runs.\r\n"
            "\r\n"
            "Another paragraph with the words Delete Me embedded.\r\n"
        )

        # wdFormatXMLDocument = 12  (*.docx)
        doc.SaveAs2(os.path.abspath(output_path), FileFormat=12)
        print(f"[OK] Test document created: {output_path}")

    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        sys.exit(1)

    finally:
        # ---------- graceful cleanup ----------
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    create_test_document()

