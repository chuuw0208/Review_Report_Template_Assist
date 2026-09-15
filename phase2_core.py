"""
Phase 2 – Core Engine: Anchor-based content migration & red-italic deletion.

This module provides the full content migration pipeline:
  1. detect_report_type()       — detect Assessment/Affirmation/Validation from filename
  2. build_section_map()        — walk headings and compute body ranges
  3. match_sections()           — pair template ↔ reference headings by normalized key
  4. migrate_sections()         — copy reference body into template (tracked changes)
  5. delete_red_italic_text()   — remove red italic instruction text (tracked changes)
  6. run_migration()            — orchestrator (called by GUI in Phase 3)

All mutations happen with TrackRevisions = True so every change appears
as a tracked revision in the output document.

Usage (CLI):
    python phase2_core.py  <template.docx>  <reference.docx>  [-o output.docx]
"""

import os
import sys
import re
import time
import argparse
import win32com.client
import pythoncom


# ============================================================================
# CONSTANTS
# ============================================================================

# Word built-in heading style names → outline level
# The bank uses custom formatting (green, etc.) but the NAMES are standard.
HEADING_STYLE_NAMES = {
    "Heading 1": 1,
    "Heading 2": 2,
    "Heading 3": 3,
}

# Word enumerations (hardcoded to avoid EnsureDispatch / MakePy)
WD_FORMAT_XML_DOC = 12     # .docx
WD_DO_NOT_SAVE    = 0      # wdDoNotSaveChanges
WD_FIND_STOP      = 0      # wdFindStop — don't wrap around
WD_COLLAPSE_END   = 0      # wdCollapseEnd

# Report type keywords (searched in filename)
REPORT_TYPES = ["Assessment", "Affirmation", "Validation"]


# ============================================================================
# HEADING NORMALISATION
# ============================================================================

def normalize_heading(text):
    """
    Normalize heading text for matching by stripping section numbers,
    APPENDIX prefixes, and normalizing whitespace/case.

    Examples:
        "1.1 Review Scope"                      → "review scope"
        "1.6.1 Model Rating and conditions..."   → "model rating and conditions..."
        "APPENDIX A: REFERENCE DOCUMENTS..."     → "reference documents..."
        "3. Governance & Controls"               → "governance & controls"
    """
    text = text.strip("\r\n\t\x07 ")

    # Strip leading section numbers: "1.", "1.1", "1.6.1", "1.6.2." etc.
    text = re.sub(r"^\d+(\.\d+)*\.?\s*", "", text)

    # Strip APPENDIX letter prefix: "APPENDIX A: " → ""
    text = re.sub(r"^APPENDIX\s+[A-Z]:\s*", "", text, flags=re.IGNORECASE)

    # Lowercase, collapse whitespace
    return " ".join(text.lower().split())


# ============================================================================
# REPORT TYPE DETECTION
# ============================================================================

def detect_report_type(filepath):
    """
    Detect report type by looking for a keyword in the filename.
    Returns: "Assessment", "Affirmation", or "Validation".
    Raises ValueError if no keyword is found.
    """
    basename = os.path.basename(filepath).lower()
    for rtype in REPORT_TYPES:
        if rtype.lower() in basename:
            return rtype
    raise ValueError(
        f"Cannot detect report type from filename '{os.path.basename(filepath)}'. "
        f"Expected one of: {REPORT_TYPES}"
    )


# ============================================================================
# SECTION MAP BUILDER
# ============================================================================

def build_section_map(doc):
    """
    Walk every paragraph in the document, find those styled as Heading 1/2/3,
    and compute the body range for each heading.

    Returns a list of dicts, one per heading:
        {
            'text':          '1.1 Review Scope',       # original text
            'key':           'review scope',            # normalised key
            'level':         2,                         # 1/2/3
            'heading_start': int,                       # Range.Start of heading para
            'heading_end':   int,                       # Range.End   of heading para
            'body_start':    int,                       # first char after heading
            'body_end':      int,                       # first char of next heading
        }

    Body range = everything from end-of-heading to start-of-next-heading
    (any level), or end-of-document for the last heading.
    """
    headings = []
    para_count = doc.Paragraphs.Count
    print(f"  [INFO] Scanning {para_count} paragraphs for headings...")

    for i in range(1, para_count + 1):
        para = doc.Paragraphs.Item(i)
        try:
            style_name = para.Style.NameLocal
        except Exception:
            continue  # skip paragraphs with corrupted/missing styles

        if style_name in HEADING_STYLE_NAMES:
            heading_text = para.Range.Text.strip("\r\n\t\x07 ")
            if not heading_text:
                continue  # skip empty headings
            headings.append({
                "text":          heading_text,
                "key":           normalize_heading(heading_text),
                "level":         HEADING_STYLE_NAMES[style_name],
                "heading_start": para.Range.Start,
                "heading_end":   para.Range.End,
            })

    # Compute body ranges
    doc_end = doc.Content.End
    for idx, h in enumerate(headings):
        h["body_start"] = h["heading_end"]
        h["body_end"] = headings[idx + 1]["heading_start"] if idx + 1 < len(headings) else doc_end

    # Pretty-print the map
    print(f"  [INFO] Found {len(headings)} heading(s):")
    for h in headings:
        indent = "  " * h["level"]
        body_chars = h["body_end"] - h["body_start"]
        print(f"  {indent}[H{h['level']}] \"{h['text']}\"  "
              f"(key: \"{h['key']}\", body: {body_chars} chars)")

    return headings


# ============================================================================
# SECTION MATCHING
# ============================================================================

def match_sections(template_headings, reference_headings):
    """
    Match template headings → reference headings by normalised key.

    Returns:
        matched           — list of (template_dict, reference_dict) tuples
        unmatched_template — template headings with no match in reference
        unmatched_reference — reference headings with no match in template
    """
    # Build reference lookup (first occurrence wins if duplicates)
    ref_lookup = {}
    for rh in reference_headings:
        ref_lookup.setdefault(rh["key"], rh)

    matched, unmatched_template = [], []
    for th in template_headings:
        ref_match = ref_lookup.get(th["key"])
        if ref_match is not None:
            matched.append((th, ref_match))
        else:
            unmatched_template.append(th)

    template_keys = {th["key"] for th in template_headings}
    unmatched_reference = [rh for rh in reference_headings if rh["key"] not in template_keys]

    # Log results
    print(f"\n  [INFO] Matching results:")
    print(f"    Matched:                {len(matched)}")
    print(f"    Template-only (kept):   {len(unmatched_template)}")
    print(f"    Reference-only (skip):  {len(unmatched_reference)}")

    if unmatched_template:
        print(f"  [NOTE] Template sections with no match (default content kept):")
        for th in unmatched_template:
            print(f"    • \"{th['text']}\"")
    if unmatched_reference:
        print(f"  [NOTE] Reference sections with no match (skipped):")
        for rh in unmatched_reference:
            print(f"    • \"{rh['text']}\"")

    return matched, unmatched_template, unmatched_reference


# ============================================================================
# CONTENT MIGRATION
# ============================================================================

def migrate_sections(template_doc, reference_doc, matched):
    """
    For each matched (template, reference) pair, INSERT the reference's
    body content into the template right after the heading — without
    deleting the template's existing pre-populated body content.

    The template's pre-populated content (black, normal font) is preserved.
    Red italic instruction text is removed separately by delete_red_italic_text().

    Processed bottom-to-top so that insertions in lower sections
    do not shift the positions of upper sections still to be processed.

    Track Changes must already be ON on template_doc.
    """
    if not matched:
        print("  [INFO] No matched sections to migrate.")
        return 0

    # Sort by template body position, descending (bottom → top)
    sorted_matches = sorted(
        matched, key=lambda m: m[0]["body_start"], reverse=True
    )

    migrated = 0
    for tmpl_h, ref_h in sorted_matches:
        ref_body_len = ref_h["body_end"] - ref_h["body_start"]

        if ref_body_len <= 1:
            print(f"  [SKIP] \"{ref_h['text']}\" — reference body is empty")
            continue

        print(f"  [MIGRATE] \"{tmpl_h['text']}\" ← ref \"{ref_h['text']}\"")

        try:
            # Step 1 — copy reference body to clipboard
            ref_body = reference_doc.Range(ref_h["body_start"], ref_h["body_end"])
            ref_body.Copy()

            # Step 2 — INSERT at body start (collapsed range = no deletion)
            # This preserves all existing template body content below.
            insert_point = template_doc.Range(
                tmpl_h["body_start"], tmpl_h["body_start"]
            )
            insert_point.Paste()

            migrated += 1
        except Exception as exc:
            print(f"  [ERROR] Failed to migrate \"{tmpl_h['text']}\": {exc}")

    print(f"\n  [INFO] Successfully migrated {migrated} section(s).")
    return migrated


# ============================================================================
# RED-ITALIC INSTRUCTION TEXT DELETION
# ============================================================================

def is_red_color(color_value):
    """
    Check whether a Word Font.Color value is 'red'.
    Word stores colours in BGR: Blue*65536 + Green*256 + Red.

    Accepts any colour with R > 180, G < 80, B < 80,
    covering pure red RGB(255,0,0)=255, dark red RGB(192,0,0)=192, etc.
    """
    if color_value is None or color_value < 0:
        return False  # wdColorAutomatic / undefined
    r = color_value & 0xFF
    g = (color_value >> 8) & 0xFF
    b = (color_value >> 16) & 0xFF
    return r > 180 and g < 80 and b < 80


def delete_red_italic_text(doc):
    """
    Find every run of italic text in the document body, check if its
    colour is red-ish, and delete it.  With TrackRevisions ON every
    deletion is recorded as a tracked change.

    Strategy:
      • Use Word Find (Format = True, Font.Italic = True) to jump from
        one italic run to the next — fast even in large documents.
      • For each match, test Font.Color with is_red_color().
      • If red → delete (tracked); restart search from doc start.
      • If not red → advance past the match and continue.

    Returns the number of deletions made.
    """
    print("\n  [INFO] Scanning for red italic instruction text...")
    delete_count = 0
    max_iter = 500  # safety cap

    search_range = doc.Range(0, doc.Content.End)

    for _ in range(max_iter):
        find = search_range.Find
        find.ClearFormatting()
        find.Font.Italic = True
        find.Text = ""           # any text with this formatting
        find.Forward = True
        find.Wrap = WD_FIND_STOP
        find.Format = True       # honour the Font criteria

        if not find.Execute():
            break  # no more italic text in the range

        # Guard: zero-length match → advance
        if search_range.Start >= search_range.End:
            break

        font_color = search_range.Font.Color
        if is_red_color(font_color):
            preview = search_range.Text[:80].replace("\r", "↵")
            print(f"  [DELETE] \"{preview}\"")
            search_range.Delete()
            delete_count += 1
            # Restart from the beginning (positions shifted)
            search_range = doc.Range(0, doc.Content.End)
        else:
            # Not red — skip forward
            next_start = search_range.End
            if next_start >= doc.Content.End:
                break
            search_range = doc.Range(next_start, doc.Content.End)

    print(f"  [INFO] Deleted {delete_count} red-italic block(s).")
    return delete_count


# ============================================================================
# MAIN ORCHESTRATOR
# ============================================================================

def run_migration(template_path, reference_path, output_path=None):
    """
    Full pipeline:
      1. Open both documents (Word invisible)
      2. Detect report type from reference filename
      3. Build section maps
      4. Match sections
      5. Migrate matched content (tracked)
      6. Delete red-italic instruction text (tracked)
      7. Save output

    This function is designed to be called from the Phase 3 GUI.
    It prints progress to stdout; a future version can accept a callback.
    """
    template_path  = os.path.abspath(template_path)
    reference_path = os.path.abspath(reference_path)
    if output_path is None:
        base, ext = os.path.splitext(template_path)
        output_path = base + "_draft_output" + ext
    output_path = os.path.abspath(output_path)

    word = None
    template_doc = None
    reference_doc = None

    try:
        pythoncom.CoInitialize()

        print("=" * 60)
        print("  Phase 2 — Content Migration Pipeline")
        print("=" * 60)

        # ---- Start Word ----
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        print("[INFO] Word started (invisible).")

        # ---- Detect report type ----
        report_type = detect_report_type(reference_path)
        print(f"[INFO] Report type detected: {report_type}\n")

        # ---- Open documents ----
        reference_doc = word.Documents.Open(reference_path, ReadOnly=True)
        print(f"[INFO] Opened reference:  {os.path.basename(reference_path)}")

        template_doc = word.Documents.Open(template_path)
        print(f"[INFO] Opened template:   {os.path.basename(template_path)}")

        # ---- Track Changes ON ----
        template_doc.TrackRevisions = True
        print("[INFO] Track Changes → ON\n")

        # ---- Build section maps ----
        print("--- Section map: Reference ---")
        ref_headings = build_section_map(reference_doc)

        print(f"\n--- Section map: Template ---")
        tmpl_headings = build_section_map(template_doc)

        # ---- Match ----
        print(f"\n--- Matching sections ---")
        matched, _, _ = match_sections(tmpl_headings, ref_headings)

        # ---- Migrate ----
        print(f"\n--- Migrating content (bottom → top) ---")
        migrated = migrate_sections(template_doc, reference_doc, matched)

        # ---- Delete red-italic ----
        print(f"\n--- Deleting red-italic instruction text ---")
        deleted = delete_red_italic_text(template_doc)

        # ---- Save ----
        if os.path.exists(output_path):
            os.remove(output_path)
        template_doc.SaveAs2(output_path, FileFormat=WD_FORMAT_XML_DOC)
        print(f"\n[INFO] Saved: {output_path}")

        # ---- Summary ----
        print("\n" + "=" * 60)
        print("  ✅  Pipeline completed successfully")
        print("=" * 60)
        print(f"  Report type      : {report_type}")
        print(f"  Sections matched : {migrated}")
        print(f"  Red-italic blocks: {deleted} deleted")
        print(f"  Output file      : {os.path.basename(output_path)}")
        print("=" * 60)
        print("\nOpen the output in Word → Review → All Markup to verify.\n")

    except Exception as exc:
        print(f"\n[ERROR] Pipeline failed: {exc}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)

    finally:
        # ---- Graceful cleanup ----
        for doc_obj in (reference_doc, template_doc):
            if doc_obj is not None:
                try:
                    doc_obj.Close(SaveChanges=WD_DO_NOT_SAVE)
                except Exception:
                    pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass

        reference_doc = None
        template_doc = None
        word = None
        pythoncom.CoUninitialize()
        time.sleep(1)
        print("[INFO] Word closed.")


# ============================================================================
# CLI
# ============================================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Migrate content from a Reference Report into a New Template "
                    "(with Track Changes)."
    )
    parser.add_argument("template",  help="Path to the New Template (.docx)")
    parser.add_argument("reference", help="Path to the Reference Report (.docx)")
    parser.add_argument("-o", "--output", help="Output file path "
                        "(default: <template>_draft_output.docx)")
    args = parser.parse_args()
    run_migration(args.template, args.reference, args.output)

