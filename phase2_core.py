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
import shutil
import argparse

try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None


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

def is_red(color_value, color_index=None):
    """
    Check whether a Word color is red.
    Word stores RGB as BGR in Font.Color: Blue*65536 + Green*256 + Red.
    Font.ColorIndex: wdRed = 6, wdDarkRed = 13.
    """
    if color_index in (6, 13):
        return True

    if color_value is None or not isinstance(color_value, int):
        return False

    if color_value < 0 or color_value == 9999999:  # wdUndefined or wdColorAutomatic
        return False

    r = color_value & 0xFF
    g = (color_value >> 8) & 0xFF
    b = (color_value >> 16) & 0xFF

    # Red must be distinctly dominant and bright enough
    return (r > 140) and (g < 110) and (b < 110) and (r > g + 40)


def _delete_red_italic_words(para_range):
    """Scan words within a mixed-formatting paragraph and delete red italic words."""
    count = 0
    try:
        words = para_range.Words
        w_count = words.Count
        for w_idx in range(w_count, 0, -1):
            try:
                w = words.Item(w_idx)
                w_font = w.Font
                if w_font.Italic in (True, -1) and is_red(w_font.Color, w_font.ColorIndex):
                    w.Delete()
                    count += 1
            except Exception:
                continue
    except Exception:
        pass
    return count


def delete_red_italic_text(doc):
    """
    Find and delete all text formatted as Red + Italic.
    With TrackRevisions ON, deletions appear as tracked changes.

    Iterates paragraphs in reverse order:
      - If an entire paragraph is Italic + Red -> delete paragraph range
      - If a paragraph has mixed formatting -> inspect words and delete red italic words
      - Explicitly protects heading styles from deletion
    """
    print("\n  [INFO] Scanning for red italic instruction text...")
    delete_count = 0
    WD_UNDEFINED = 9999999

    para_count = doc.Paragraphs.Count
    print(f"  [INFO] Checking {para_count} paragraphs...")

    # Iterate backwards so deletions don't shift indices of remaining paragraphs
    for i in range(para_count, 0, -1):
        try:
            para = doc.Paragraphs.Item(i)
        except Exception:
            continue

        # Never delete headings
        try:
            if para.Style.NameLocal in HEADING_STYLE_NAMES:
                continue
        except Exception:
            pass

        rng = para.Range
        font = rng.Font

        try:
            italic_val = font.Italic
        except Exception:
            continue

        # Fast path: Paragraph is not italic at all
        if italic_val == 0 or italic_val is False:
            continue

        # Fast path: Entire paragraph is italic
        if italic_val in (True, -1):
            try:
                color_val = font.Color
                color_idx = font.ColorIndex
            except Exception:
                color_val = None
                color_idx = None

            if is_red(color_val, color_idx):
                preview = rng.Text[:60].replace("\r", " ").strip()
                if preview:
                    print(f"  [DELETE] \"{preview}\"")
                rng.Delete()
                delete_count += 1
                continue
            elif color_val == WD_UNDEFINED:
                # Mixed colors within italic paragraph
                delete_count += _delete_red_italic_words(rng)
                continue

        # Mixed italic formatting within paragraph
        if italic_val == WD_UNDEFINED:
            delete_count += _delete_red_italic_words(rng)

    print(f"  [INFO] Deleted {delete_count} red-italic instruction block(s).")
    return delete_count


def remove_affirmation_excluded_sections(template_doc):
    """
    For Affirmation reports, eliminate Section 2 ('Additional Analysis since prior review'
    and subsections 2.1, 2.2, 2.3).
    With TrackRevisions ON, this deletion is recorded as tracked changes.
    """
    headings = build_section_map(template_doc)
    for idx, h in enumerate(headings):
        if "additional analysis since prior review" in h["key"]:
            start_pos = h["heading_start"]
            end_pos = h["body_end"]
            for next_h in headings[idx + 1:]:
                if next_h["level"] == 1:
                    end_pos = next_h["heading_start"]
                    break
                else:
                    end_pos = next_h["body_end"]

            print(f"\n  [AFFIRMATION] Removing Section 2 ('{h['text']}' and subsections) under Track Changes...")
            del_range = template_doc.Range(start_pos, end_pos)
            del_range.Delete()
            return True
    return False


def extract_cover_title(doc, fallback_name="Report"):
    """
    Extract the title from the cover page (title page) of the document.
    Scans paragraphs prior to the first heading (Heading 1/2/3).
    Selects prominent title text (styled as 'Title' or bold/large text).
    Sanitizes the extracted string for a valid Windows filename.
    """
    title_candidates = []
    para_count = doc.Paragraphs.Count

    for i in range(1, min(para_count + 1, 25)):
        try:
            para = doc.Paragraphs.Item(i)
            style_name = para.Style.NameLocal
        except Exception:
            continue

        # Stop scanning when we hit the first standard heading (e.g. "1. Review Details")
        if style_name in HEADING_STYLE_NAMES:
            break

        text = para.Range.Text.strip("\r\n\t\x07 ")
        if not text:
            continue

        # Check if styled as 'Title'
        if "title" in style_name.lower() and "subtitle" not in style_name.lower():
            title_candidates.append(text)
            break

        # Check if paragraph has bold font or prominent size (>= 14pt)
        font = para.Range.Font
        try:
            is_bold = (font.Bold in (True, -1))
            size = font.Size
        except Exception:
            is_bold = False
            size = 11

        lower_t = text.lower()
        if any(prefix in lower_t for prefix in ["version", "ver.", "date:", "author:", "prepared by:", "model id:"]):
            continue

        if is_bold or (size and size >= 14):
            title_candidates.append(text)

    if title_candidates:
        raw_title = " ".join(title_candidates)
    else:
        try:
            prop_title = doc.BuiltInDocumentProperties("Title").Value.strip()
            if prop_title and len(prop_title) > 2:
                raw_title = prop_title
            else:
                raw_title = fallback_name
        except Exception:
            raw_title = fallback_name

    # Sanitize for valid Windows filename: remove < > : " / \ | ? *
    clean_title = re.sub(r'[\\/*?:"<>|\r\n\t]+', ' ', raw_title)
    clean_title = " ".join(clean_title.split()).strip()

    if len(clean_title) > 80:
        clean_title = clean_title[:80].strip()

    return clean_title if clean_title else fallback_name


# ============================================================================
# MAIN ORCHESTRATOR
# ============================================================================

def run_migration(template_path, reference_path, output_path=None, target_report_type="Assessment", progress_callback=None):
    """
    Full pipeline:
      1. Open Word (invisible)
      2. If output_path is None, extract cover page title from template and name output as:
         <Cover Page Title>_Draft.docx
      3. Create clean working copy of template at output_path (template is untouched!)
      4. Detect report type from reference filename
      5. Build section maps & match sections
      6. Migrate matched content (tracked)
      7. Delete red-italic instruction text (tracked)
      8. If target is Affirmation, prune Section 2 under Track Changes
      9. Save output and return output_path
    """
    def report_progress(percent, text):
        if progress_callback:
            try:
                progress_callback(percent, text)
            except Exception:
                pass

    template_path  = os.path.abspath(template_path)
    reference_path = os.path.abspath(reference_path)

    if not os.path.isfile(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")
    if not os.path.isfile(reference_path):
        raise FileNotFoundError(f"Reference report not found: {reference_path}")

    word = None
    template_doc = None
    reference_doc = None

    try:
        pythoncom.CoInitialize()

        print("=" * 60)
        print("  Phase 2 — Content Migration Pipeline")
        print("=" * 60)

        report_progress(10, "Starting Word application...")

        # ---- Start Word ----
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        print("[INFO] Word started (invisible).")

        # ---- Determine output_path from cover page title if None ----
        if output_path is None:
            report_progress(15, "Detecting cover page title from template...")
            temp_tmpl = word.Documents.Open(template_path, ReadOnly=True)
            tmpl_dir = os.path.dirname(template_path)
            tmpl_base, ext = os.path.splitext(os.path.basename(template_path))
            cover_title = extract_cover_title(temp_tmpl, fallback_name=tmpl_base)
            temp_tmpl.Close(SaveChanges=WD_DO_NOT_SAVE)
            temp_tmpl = None

            output_filename = f"{cover_title}_Draft{ext}"
            output_path = os.path.join(tmpl_dir, output_filename)
            print(f"[INFO] Cover page title: \"{cover_title}\" -> Draft filename: \"{output_filename}\"")
        else:
            output_path = os.path.abspath(output_path)

        report_progress(20, f"Creating working copy -> {os.path.basename(output_path)}...")

        # Guard: Ensure output file is not locked by an existing open Word window
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except PermissionError:
                raise RuntimeError(
                    f"Cannot write to '{os.path.basename(output_path)}'. "
                    f"Please close the document in Word and run the script again."
                )

        # Copy template to output_path at the OS level BEFORE opening for modification
        shutil.copy2(template_path, output_path)

        # ---- Detect report type ----
        report_type = detect_report_type(reference_path)
        print(f"[INFO] Report type detected: {report_type}\n")

        report_progress(35, "Opening documents and mapping sections...")

        # ---- Open documents ----
        reference_doc = word.Documents.Open(reference_path, ReadOnly=True)
        print(f"[INFO] Opened reference (read-only): {os.path.basename(reference_path)}")

        # Open working copy at output_path. Original template_path is NEVER opened for writing.
        template_doc = word.Documents.Open(output_path)
        print(f"[INFO] Opened working copy:          {os.path.basename(output_path)}")

        # ---- Track Changes ON ----
        template_doc.TrackRevisions = True
        print("[INFO] Track Changes -> ON\n")

        # ---- Build section maps ----
        print("--- Section map: Reference ---")
        ref_headings = build_section_map(reference_doc)

        print(f"\n--- Section map: Template ---")
        tmpl_headings = build_section_map(template_doc)

        # ---- Match ----
        print(f"\n--- Matching sections ---")
        matched, _, _ = match_sections(tmpl_headings, ref_headings)

        report_progress(60, f"Migrating {len(matched)} matched sections...")

        # ---- Migrate ----
        print(f"\n--- Migrating content (bottom -> top) ---")
        migrated = migrate_sections(template_doc, reference_doc, matched)

        report_progress(80, "Scanning and deleting red-italic instructions...")

        # ---- Delete red-italic ----
        print(f"\n--- Deleting red-italic instruction text ---")
        deleted = delete_red_italic_text(template_doc)

        # ---- Target Report Type tailoring (Affirmation pruning) ----
        if str(target_report_type).strip().lower() == "affirmation":
            report_progress(90, "Pruning Section 2 for Affirmation report...")
            print(f"\n--- Tailoring for Affirmation Report ---")
            remove_affirmation_excluded_sections(template_doc)

        report_progress(95, "Saving output draft document...")

        # ---- Save ----
        template_doc.Save()
        print(f"\n[INFO] Saved output document: {output_path}")

        report_progress(100, "Draft report generated successfully!")

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
        return output_path

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

