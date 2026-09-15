"""
Phase 1 – PoC: Prove pywin32 COM automation works in the bank IT environment.

This script:
  1. Opens 'test_input.docx' invisibly in the background.
  2. Turns ON Track Changes (TrackRevisions).
  3. Finds every occurrence of "Delete Me" and deletes it (recorded as a
     tracked deletion — red strikethrough in Word).
  4. Inserts "Hello World" at the end of the document (recorded as a
     tracked insertion — underlined text in Word).
  5. Saves the result as 'test_output.docx' and closes Word gracefully.

Usage:
    python phase1_poc.py

Pre-requisite:
    Run  phase1_create_test_doc.py  first to generate test_input.docx,
    OR point INPUT_FILENAME below at any existing .docx you like.
"""

import os
import sys
import time
import win32com.client
import pythoncom

# ---------- configuration ----------
SCRIPT_DIR       = os.path.dirname(os.path.abspath(__file__))
INPUT_FILENAME   = "test_input.docx"
OUTPUT_FILENAME  = "test_output.docx"
SEARCH_STRING    = "Delete Me"
INSERT_STRING    = "Hello World"

# Word enumerations (avoiding the need for win32com.client.constants / EnsureDispatch)
WD_REPLACE_ALL       = 2      # wdReplaceAll
WD_FORMAT_XML_DOC    = 12     # wdFormatXMLDocument  (.docx)
WD_DO_NOT_SAVE       = 0      # wdDoNotSaveChanges
WD_FIND_CONTINUE     = 1      # wdFindContinue
WD_STORY_MAIN        = 6      # wdMainTextStory (not needed but documented)
WD_COLLAPSE_END      = 0      # wdCollapseEnd


def run_poc():
    """Main PoC routine."""
    input_path  = os.path.join(SCRIPT_DIR, INPUT_FILENAME)
    output_path = os.path.join(SCRIPT_DIR, OUTPUT_FILENAME)

    if not os.path.isfile(input_path):
        print(f"[ERROR] Input file not found: {input_path}", file=sys.stderr)
        print("        Run  phase1_create_test_doc.py  first.", file=sys.stderr)
        sys.exit(1)

    word = None
    doc  = None

    try:
        # ---- COM initialisation ----
        pythoncom.CoInitialize()

        # DispatchEx starts a *new* Word instance (avoids hijacking user session)
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False   # no "do you want to save?" popups

        print(f"[INFO] Word application started (PID will be released on Quit).")

        # ---- Step 1: Open the document ----
        doc = word.Documents.Open(os.path.abspath(input_path))
        print(f"[INFO] Opened: {input_path}")

        # ---- Step 2: Enable Track Changes ----
        doc.TrackRevisions = True
        print("[INFO] Track Changes → ON")

        # ---- Step 3: Find & Delete "Delete Me" via tracked deletion ----
        #
        #   With TrackRevisions ON, Range.Delete() records a tracked deletion.
        #   We use Find.Execute in a loop so we can delete the Range each time,
        #   which gives us explicit control (vs. ReplaceAll which can silently
        #   skip certain edge cases in complex documents).
        #
        delete_count = 0
        find_range = doc.Content

        while True:
            find = find_range.Find
            find.ClearFormatting()
            find.Text = SEARCH_STRING
            find.Forward = True
            find.Wrap = WD_FIND_CONTINUE
            find.MatchCase = False
            find.MatchWholeWord = False

            if not find.Execute():
                break  # no more matches

            # find_range is now collapsed to the found text
            # Delete the found text — this records a tracked deletion
            find_range.Delete()
            delete_count += 1

            # Reset the search range to the rest of the document to avoid
            # infinite loop (after deletion the range collapses to a point)
            find_range = doc.Content

        print(f"[INFO] Deleted {delete_count} occurrence(s) of \"{SEARCH_STRING}\" "
              f"(tracked as revisions).")

        # ---- Step 4: Insert text at the end of the document ----
        end_range = doc.Content
        end_range.Collapse(Direction=WD_COLLAPSE_END)   # move to very end
        end_range.InsertParagraphAfter()
        end_range.Collapse(Direction=WD_COLLAPSE_END)
        end_range.InsertAfter(INSERT_STRING)
        print(f"[INFO] Inserted \"{INSERT_STRING}\" at end of document (tracked).")

        # ---- Step 5: Save As new file (.docx) ----
        if os.path.exists(output_path):
            os.remove(output_path)          # avoid "file already exists" dialog
        doc.SaveAs2(os.path.abspath(output_path), FileFormat=WD_FORMAT_XML_DOC)
        print(f"[INFO] Saved: {output_path}")

        # ---- SUCCESS ----
        print("\n========================================")
        print(" Phase 1 PoC PASSED")
        print("========================================")
        print(f" Open '{OUTPUT_FILENAME}' in Word and verify:")
        print(f"   • \"{SEARCH_STRING}\" appears as red strikethrough (tracked deletion)")
        print(f"   • \"{INSERT_STRING}\" appears as underlined text (tracked insertion)")
        print("========================================\n")

    except Exception as exc:
        print(f"\n[ERROR] PoC FAILED: {exc}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)

    finally:
        # ---- Graceful cleanup — CRITICAL to avoid ghost WINWORD.EXE ----
        if doc is not None:
            try:
                doc.Close(SaveChanges=WD_DO_NOT_SAVE)
            except Exception:
                pass

        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass

        # Release COM references explicitly, then give Windows a moment
        doc  = None
        word = None
        pythoncom.CoUninitialize()

        # Small delay to let the OS fully release the process
        time.sleep(1)
        print("[INFO] Word application closed. No ghost processes should remain.")


if __name__ == "__main__":
    run_poc()

