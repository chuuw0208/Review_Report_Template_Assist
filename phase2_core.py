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
import copy
import zipfile
import datetime
import xml.etree.ElementTree as ET

try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None

# OpenXML Namespaces
OPENXML_NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v': 'urn:schemas-microsoft-com:vml',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
}
for _p, _u in OPENXML_NS.items():
    ET.register_namespace(_p, _u)

W_NS = OPENXML_NS['w']
def _w_tag(name):
    return f"{{{W_NS}}}{name}"

_P_TAG = _w_tag("p")
_R_TAG = _w_tag("r")
_T_TAG = _w_tag("t")
_PPR_TAG = _w_tag("pPr")
_PSTYLE_TAG = _w_tag("pStyle")
_RPR_TAG = _w_tag("rPr")
_I_TAG = _w_tag("i")
_B_TAG = _w_tag("b")
_SZ_TAG = _w_tag("sz")
_COLOR_TAG = _w_tag("color")
_INS_TAG = _w_tag("ins")
_DEL_TAG = _w_tag("del")
_DELTEXT_TAG = _w_tag("delText")
_BODY_TAG = _w_tag("body")
_TBL_TAG = _w_tag("tbl")
_SETTINGS_TAG = _w_tag("settings")
_TRACK_REV_TAG = _w_tag("trackRevisions")

REVISION_AUTHOR = "MRMV Validator"



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
    Falls back to "Assessment" if no keyword is found.
    """
    basename = os.path.basename(filepath).lower()
    for rtype in REPORT_TYPES:
        if rtype.lower() in basename:
            return rtype
    return "Assessment"



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
    For Affirmation reports:
      1. Eliminate Section 2 ('Additional Analysis since prior review' and subsections 2.1, 2.2, 2.3).
      2. Renumber Section 3 ('3. Governance & Controls' and 3.1, 3.2, etc.) to Section 2 ('2.', '2.1', '2.2').
    With TrackRevisions ON, deletions and renumbering appear as tracked revisions.
    """
    headings = build_section_map(template_doc)
    deleted_sec2 = False

    # Step 1: Delete Section 2
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
            deleted_sec2 = True
            break

    # Step 2: Renumber Section 3 -> Section 2 (3. -> 2., 3.1 -> 2.1, 3.2 -> 2.2)
    print(f"\n  [AFFIRMATION] Renumbering Section 3 to Section 2 under Track Changes...")
    para_count = template_doc.Paragraphs.Count
    for i in range(1, para_count + 1):
        try:
            para = template_doc.Paragraphs.Item(i)
            style_name = para.Style.NameLocal
        except Exception:
            continue

        if style_name in HEADING_STYLE_NAMES:
            text = para.Range.Text
            # Look for leading "3." in heading
            m = re.match(r'^(\s*)3\.', text)
            if m:
                leading_spaces = len(m.group(1))
                num_start = para.Range.Start + leading_spaces
                num_end = num_start + 2  # len of "3." is 2
                num_range = template_doc.Range(num_start, num_end)
                if num_range.Text == "3.":
                    print(f"  [RENUMBER] '{text.strip()}' -> '2.{text.strip()[leading_spaces+2:]}'")
                    num_range.Text = "2."
                else:
                    try:
                        f = para.Range.Find
                        f.ClearFormatting()
                        f.Text = "3."
                        f.Replacement.ClearFormatting()
                        f.Replacement.Text = "2."
                        f.Forward = True
                        f.Wrap = 0  # wdFindStop
                        f.Execute(Replace=1)  # wdReplaceOne
                        print(f"  [RENUMBER via Find] '{text.strip()}'")
                    except Exception as e:
                        print(f"  [WARNING] Renumber failed for '{text.strip()}': {e}")

    return deleted_sec2


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
# WINDOWS WORD COM ENGINE
# ============================================================================

def run_migration_windows_com(template_path, reference_path, output_path=None, target_report_type="Assessment", progress_callback=None):
    """
    Windows-native Word COM pipeline (uses installed MS Word application).
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
        print("  Phase 2 — Content Migration Pipeline (Windows COM)")
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
        raise exc

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
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        time.sleep(1)
        print("[INFO] Word closed.")


# ============================================================================
# CROSS-PLATFORM OPENXML ENGINE (MACOS / LINUX / ZERO DEPENDENCIES)
# ============================================================================

def _get_iso_now():
    return datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

def _is_red_hex(val):
    if not val:
        return False
    val = val.strip().lower()
    if val in ("red", "ff0000", "c00000", "ed1c24", "ee2726", "d00000"):
        return True
    if len(val) == 6:
        try:
            r = int(val[0:2], 16)
            g = int(val[2:4], 16)
            b = int(val[4:6], 16)
            return (r > 140) and (g < 110) and (b < 110) and (r > g + 40)
        except ValueError:
            pass
    return False

def _is_italic_run(r_elem):
    rPr = r_elem.find(_RPR_TAG)
    if rPr is not None:
        i_elem = rPr.find(_I_TAG)
        if i_elem is not None:
            val = i_elem.attrib.get(_w_tag("val"), "true").lower()
            return val not in ("false", "0")
    return False

def _is_red_run(r_elem):
    rPr = r_elem.find(_RPR_TAG)
    if rPr is not None:
        c_elem = rPr.find(_COLOR_TAG)
        if c_elem is not None:
            val = c_elem.attrib.get(_w_tag("val"), "")
            return _is_red_hex(val)
    return False

def _get_elem_text(elem):
    return "".join(elem.itertext()).strip()

def get_heading_info_openxml(p_elem):
    pPr = p_elem.find(_PPR_TAG)
    if pPr is not None:
        pStyle = pPr.find(_PSTYLE_TAG)
        if pStyle is not None:
            val = pStyle.attrib.get(_w_tag("val"), "").lower()
            m = re.search(r'heading\s*([1-3])', val)
            if m:
                level = int(m.group(1))
                text = _get_elem_text(p_elem)
                return level, text
        outline = pPr.find(_w_tag("outlineLvl"))
        if outline is not None:
            try:
                lvl_val = int(outline.attrib.get(_w_tag("val"), "-1"))
                if 0 <= lvl_val <= 2:
                    return lvl_val + 1, _get_elem_text(p_elem)
            except ValueError:
                pass
    return None, None

def parse_docx_sections_openxml(body_elem):
    children = list(body_elem)
    sections = []
    current_sec = None

    for idx, child in enumerate(children):
        if child.tag == _P_TAG:
            lvl, text = get_heading_info_openxml(child)
            if lvl is not None and text:
                current_sec = {
                    'level': lvl,
                    'text': text,
                    'key': normalize_heading(text),
                    'heading_index': idx,
                    'heading_element': child,
                    'body_elements': []
                }
                sections.append(current_sec)
                continue

        if current_sec is not None:
            current_sec['body_elements'].append(child)

    return sections

def wrap_in_tracked_ins_openxml(elem, rev_id_gen):
    now_str = _get_iso_now()
    if elem.tag == _P_TAG:
        new_children = []
        for child in list(elem):
            if child.tag == _PPR_TAG:
                rPr = child.find(_RPR_TAG)
                if rPr is None:
                    rPr = ET.SubElement(child, _RPR_TAG)
                ins_pr = ET.SubElement(rPr, _INS_TAG)
                ins_pr.attrib[_w_tag("id")] = str(rev_id_gen[0])
                ins_pr.attrib[_w_tag("author")] = REVISION_AUTHOR
                ins_pr.attrib[_w_tag("date")] = now_str
                rev_id_gen[0] += 1
                new_children.append(child)
            elif child.tag == _R_TAG:
                ins = ET.Element(_INS_TAG)
                ins.attrib[_w_tag("id")] = str(rev_id_gen[0])
                ins.attrib[_w_tag("author")] = REVISION_AUTHOR
                ins.attrib[_w_tag("date")] = now_str
                rev_id_gen[0] += 1
                ins.append(child)
                new_children.append(ins)
            elif child.tag == _INS_TAG:
                new_children.append(child)
            else:
                new_children.append(child)
        elem[:] = new_children

    elif elem.tag == _TBL_TAG:
        for p in elem.iter(_P_TAG):
            wrap_in_tracked_ins_openxml(p, rev_id_gen)

def delete_element_tracked_openxml(elem, rev_id_gen):
    now_str = _get_iso_now()
    if elem.tag == _P_TAG:
        pPr = elem.find(_PPR_TAG)
        if pPr is not None:
            rPr = pPr.find(_RPR_TAG)
            if rPr is None:
                rPr = ET.SubElement(pPr, _RPR_TAG)
            del_pr = ET.SubElement(rPr, _DEL_TAG)
            del_pr.attrib[_w_tag("id")] = str(rev_id_gen[0])
            del_pr.attrib[_w_tag("author")] = REVISION_AUTHOR
            del_pr.attrib[_w_tag("date")] = now_str
            rev_id_gen[0] += 1

        new_children = []
        for child in list(elem):
            if child.tag == _PPR_TAG:
                new_children.append(child)
            elif child.tag == _R_TAG:
                del_elem = ET.Element(_DEL_TAG)
                del_elem.attrib[_w_tag("id")] = str(rev_id_gen[0])
                del_elem.attrib[_w_tag("author")] = REVISION_AUTHOR
                del_elem.attrib[_w_tag("date")] = now_str
                rev_id_gen[0] += 1

                del_r = ET.SubElement(del_elem, _R_TAG)
                rPr = child.find(_RPR_TAG)
                if rPr is not None:
                    del_r.append(copy.deepcopy(rPr))
                for t in child.findall(_T_TAG):
                    del_t = ET.SubElement(del_r, _DELTEXT_TAG)
                    del_t.text = t.text
                    del_t.attrib["{http://www.w3.org/XML/1998/namespace}space"] = "preserve"
                new_children.append(del_elem)
            elif child.tag == _DEL_TAG:
                new_children.append(child)
            else:
                new_children.append(child)
        elem[:] = new_children
    elif elem.tag == _TBL_TAG:
        for p in elem.iter(_P_TAG):
            delete_element_tracked_openxml(p, rev_id_gen)

def delete_red_italic_openxml(body_elem, rev_id_gen):
    now_str = _get_iso_now()
    deleted_count = 0

    for p in body_elem.findall(_P_TAG):
        runs = p.findall(_R_TAG)
        if not runs:
            continue

        all_red_italic = True
        has_text = False
        for r in runs:
            t = "".join(r.itertext()).strip()
            if t:
                has_text = True
                if not (_is_italic_run(r) and _is_red_run(r)):
                    all_red_italic = False
                    break

        if has_text and all_red_italic:
            delete_element_tracked_openxml(p, rev_id_gen)
            deleted_count += 1
            continue

        new_children = []
        p_modified = False
        for child in list(p):
            if child.tag == _R_TAG and _is_italic_run(child) and _is_red_run(child):
                del_elem = ET.Element(_DEL_TAG)
                del_elem.attrib[_w_tag("id")] = str(rev_id_gen[0])
                del_elem.attrib[_w_tag("author")] = REVISION_AUTHOR
                del_elem.attrib[_w_tag("date")] = now_str
                rev_id_gen[0] += 1

                del_r = ET.SubElement(del_elem, _R_TAG)
                rPr = child.find(_RPR_TAG)
                if rPr is not None:
                    del_r.append(copy.deepcopy(rPr))
                for t in child.findall(_T_TAG):
                    del_t = ET.SubElement(del_r, _DELTEXT_TAG)
                    del_t.text = t.text
                    del_t.attrib["{http://www.w3.org/XML/1998/namespace}space"] = "preserve"

                new_children.append(del_elem)
                deleted_count += 1
                p_modified = True
            else:
                new_children.append(child)

        if p_modified:
            p[:] = new_children

    return deleted_count

def tailor_affirmation_openxml(body_elem, rev_id_gen):
    now_str = _get_iso_now()
    children = list(body_elem)

    # 1. Delete Section 2 and all subsections
    sec2_start_idx = None
    sec2_end_idx = len(children)
    for idx, child in enumerate(children):
        if child.tag == _P_TAG:
            lvl, text = get_heading_info_openxml(child)
            if lvl == 1:
                norm = normalize_heading(text)
                if "additional analysis since prior review" in norm:
                    sec2_start_idx = idx
                elif sec2_start_idx is not None and idx > sec2_start_idx:
                    sec2_end_idx = idx
                    break

    if sec2_start_idx is not None:
        print(f"  [AFFIRMATION] Deleting Section 2 and all subsections (elements {sec2_start_idx} to {sec2_end_idx})...")
        for i in range(sec2_start_idx, sec2_end_idx):
            delete_element_tracked_openxml(children[i], rev_id_gen)

    # 2. Renumber Section 3 -> Section 2
    renum_count = 0
    for child in children:
        if child.tag == _P_TAG:
            lvl, text = get_heading_info_openxml(child)
            if lvl is not None and text:
                m = re.match(r'^(\s*)(3)(\.|\.\d+)', text)
                if m:
                    for r in child.findall(_R_TAG):
                        for t in r.findall(_T_TAG):
                            if t.text and re.match(r'^(\s*)3(\.|\.\d+)', t.text):
                                old_str = t.text
                                m_run = re.match(r'^(\s*)(3)(\.|\.\d+)(.*)', old_str)
                                if m_run:
                                    leading = m_run.group(1)
                                    num_tail = m_run.group(3)
                                    rest = m_run.group(4)

                                    old_num = f"{leading}3{num_tail}"
                                    new_num = f"{leading}2{num_tail}"

                                    r_idx = list(child).index(r)
                                    child.remove(r)

                                    # Tracked del old
                                    del_e = ET.Element(_DEL_TAG)
                                    del_e.attrib[_w_tag("id")] = str(rev_id_gen[0])
                                    del_e.attrib[_w_tag("author")] = REVISION_AUTHOR
                                    del_e.attrib[_w_tag("date")] = now_str
                                    rev_id_gen[0] += 1
                                    del_r = ET.SubElement(del_e, _R_TAG)
                                    del_t = ET.SubElement(del_r, _DELTEXT_TAG)
                                    del_t.text = old_num
                                    del_t.attrib["{http://www.w3.org/XML/1998/namespace}space"] = "preserve"

                                    # Tracked ins new
                                    ins_e = ET.Element(_INS_TAG)
                                    ins_e.attrib[_w_tag("id")] = str(rev_id_gen[0])
                                    ins_e.attrib[_w_tag("author")] = REVISION_AUTHOR
                                    ins_e.attrib[_w_tag("date")] = now_str
                                    rev_id_gen[0] += 1
                                    ins_r = ET.SubElement(ins_e, _R_TAG)
                                    ins_t = ET.SubElement(ins_r, _T_TAG)
                                    ins_t.text = new_num
                                    ins_t.attrib["{http://www.w3.org/XML/1998/namespace}space"] = "preserve"

                                    rest_r = None
                                    if rest:
                                        rest_r = ET.Element(_R_TAG)
                                        rest_t = ET.SubElement(rest_r, _T_TAG)
                                        rest_t.text = rest
                                        rest_t.attrib["{http://www.w3.org/XML/1998/namespace}space"] = "preserve"

                                    child.insert(r_idx, del_e)
                                    child.insert(r_idx + 1, ins_e)
                                    if rest_r is not None:
                                        child.insert(r_idx + 2, rest_r)

                                    renum_count += 1
                                    print(f"  [AFFIRMATION] Renumbered '{old_str.strip()}' -> '{new_num}{rest}'")
                                    break
    return renum_count

def extract_cover_title_openxml(tree, fallback_name="Report"):
    root = tree.getroot()
    body = root.find(_BODY_TAG)
    if body is None:
        return fallback_name

    title_candidates = []
    for p in body.findall(_P_TAG):
        pPr = p.find(_PPR_TAG)
        if pPr is not None:
            pStyle = pPr.find(_PSTYLE_TAG)
            if pStyle is not None:
                sval = pStyle.attrib.get(_w_tag("val"), "").lower()
                if "heading" in sval:
                    break
                if "title" in sval and "subtitle" not in sval:
                    t = "".join(p.itertext()).strip()
                    if t:
                        title_candidates.append(t)
                        break

        t = "".join(p.itertext()).strip()
        if not t:
            continue
        lower_t = t.lower()
        if any(prefix in lower_t for prefix in ["version", "ver.", "date:", "author:", "prepared by:", "model id:"]):
            continue

        for r in p.findall(_R_TAG):
            rPr = r.find(_RPR_TAG)
            if rPr is not None:
                b = rPr.find(_B_TAG)
                sz = rPr.find(_SZ_TAG)
                is_bold = (b is not None and b.attrib.get(_w_tag("val"), "true") not in ("false", "0"))
                sz_val = 0
                if sz is not None:
                    try:
                        sz_val = int(sz.attrib.get(_w_tag("val"), "0"))
                    except ValueError:
                        pass
                if is_bold or sz_val >= 28:
                    title_candidates.append(t)
                    break
        if title_candidates:
            break

    if title_candidates:
        raw_title = " ".join(title_candidates)
    else:
        raw_title = fallback_name

    clean_title = re.sub(r'[\\/*?:"<>|\r\n\t]+', ' ', raw_title)
    clean_title = " ".join(clean_title.split()).strip()
    if len(clean_title) > 80:
        clean_title = clean_title[:80].strip()
    return clean_title if clean_title else fallback_name

def run_migration_openxml(template_path, reference_path, output_path=None, target_report_type="Assessment", progress_callback=None):
    """
    Pure-Python OpenXML pipeline. Runs anywhere without Word or external dependencies.
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

    print("=" * 60)
    print("  Phase 2 — Content Migration Pipeline (OpenXML)")
    print("=" * 60)

    report_progress(10, "Reading template and reference packages...")

    rev_id_gen = [1000]

    with zipfile.ZipFile(template_path, "r") as z_tmpl:
        tmpl_files = {name: z_tmpl.read(name) for name in z_tmpl.namelist()}

    with zipfile.ZipFile(reference_path, "r") as z_ref:
        ref_files = {name: z_ref.read(name) for name in z_ref.namelist()}

    report_progress(20, "Detecting cover title and configuring output path...")

    tmpl_tree = ET.ElementTree(ET.fromstring(tmpl_files["word/document.xml"]))
    tmpl_root = tmpl_tree.getroot()
    tmpl_body = tmpl_root.find(_BODY_TAG)

    if output_path is None:
        tmpl_dir = os.path.dirname(template_path)
        tmpl_base, ext = os.path.splitext(os.path.basename(template_path))
        cover_title = extract_cover_title_openxml(tmpl_tree, fallback_name=tmpl_base)
        output_filename = f"{cover_title}_Draft{ext}"
        output_path = os.path.join(tmpl_dir, output_filename)
        print(f"[INFO] Cover page title: \"{cover_title}\" -> Draft filename: \"{output_filename}\"")
    else:
        output_path = os.path.abspath(output_path)

    report_progress(35, "Parsing section structures...")

    ref_tree = ET.ElementTree(ET.fromstring(ref_files["word/document.xml"]))
    ref_root = ref_tree.getroot()
    ref_body = ref_root.find(_BODY_TAG)

    tmpl_sections = parse_docx_sections_openxml(tmpl_body)
    ref_sections = parse_docx_sections_openxml(ref_body)
    ref_sec_by_key = {s['key']: s for s in ref_sections}

    report_progress(50, "Migrating matched sections with Track Changes...")

    migrated_count = 0
    for s in reversed(tmpl_sections):
        k = s['key']
        if k in ref_sec_by_key:
            ref_s = ref_sec_by_key[k]
            if ref_s['body_elements']:
                h_elem = s['heading_element']
                h_idx = list(tmpl_body).index(h_elem)
                insert_pos = h_idx + 1
                for ref_elem in ref_s['body_elements']:
                    clone = ET.fromstring(ET.tostring(ref_elem))
                    wrap_in_tracked_ins_openxml(clone, rev_id_gen)
                    tmpl_body.insert(insert_pos, clone)
                    insert_pos += 1
                migrated_count += 1
                print(f"  [MIGRATE] Migrated: {s['text']}")

    report_progress(75, "Scanning and deleting red-italic instructions...")
    del_inst_count = delete_red_italic_openxml(tmpl_body, rev_id_gen)
    print(f"  [INFO] Deleted {del_inst_count} red-italic instruction block(s).")

    if str(target_report_type).strip().lower() == "affirmation":
        report_progress(85, "Tailoring Affirmation report (pruning Section 2 & renumbering)...")
        tailor_affirmation_openxml(tmpl_body, rev_id_gen)

    report_progress(92, "Enabling Track Changes in settings...")

    settings_data = tmpl_files.get("word/settings.xml", None)
    if settings_data is not None:
        s_tree = ET.ElementTree(ET.fromstring(settings_data))
        s_root = s_tree.getroot()
        if s_root.find(_TRACK_REV_TAG) is None:
            ET.SubElement(s_root, _TRACK_REV_TAG)
        tmpl_files["word/settings.xml"] = ET.tostring(s_root, encoding="utf-8", xml_declaration=True)
    else:
        new_settings = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="{W_NS}"><w:trackRevisions/></w:settings>'
        tmpl_files["word/settings.xml"] = new_settings.encode("utf-8")

    tmpl_files["word/document.xml"] = ET.tostring(tmpl_root, encoding="utf-8", xml_declaration=True)

    for r_name, r_bytes in ref_files.items():
        if r_name.startswith("word/media/") and r_name not in tmpl_files:
            tmpl_files[r_name] = r_bytes

    report_progress(96, f"Writing output package -> {os.path.basename(output_path)}...")

    if os.path.exists(output_path):
        try:
            os.remove(output_path)
        except Exception:
            pass

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as z_out:
        for fname, content in tmpl_files.items():
            z_out.writestr(fname, content)

    report_progress(100, "Draft report generated successfully!")

    print("\n" + "=" * 60)
    print("  ✅  Pipeline completed successfully (OpenXML)")
    print("=" * 60)
    print(f"  Report type      : {detect_report_type(reference_path)}")
    print(f"  Sections matched : {migrated_count}")
    print(f"  Red-italic blocks: {del_inst_count} deleted")
    print(f"  Output file      : {os.path.basename(output_path)}")
    print("=" * 60)
    return output_path


# ============================================================================
# UNIFIED CROSS-PLATFORM DISPATCHER
# ============================================================================

def run_migration(template_path, reference_path, output_path=None, target_report_type="Assessment", progress_callback=None):
    """
    Cross-Platform Entry Point:
      - On Windows with pywin32: uses Word COM automation (run_migration_windows_com)
      - On macOS, Linux, or without pywin32: uses pure-Python OpenXML engine (run_migration_openxml)
    """
    if sys.platform == "win32" and win32com is not None:
        try:
            return run_migration_windows_com(
                template_path=template_path,
                reference_path=reference_path,
                output_path=output_path,
                target_report_type=target_report_type,
                progress_callback=progress_callback
            )
        except Exception as exc:
            print(f"[WARNING] Windows COM engine encountered an issue: {exc}. Falling back to OpenXML engine.")
            return run_migration_openxml(
                template_path=template_path,
                reference_path=reference_path,
                output_path=output_path,
                target_report_type=target_report_type,
                progress_callback=progress_callback
            )
    else:
        return run_migration_openxml(
            template_path=template_path,
            reference_path=reference_path,
            output_path=output_path,
            target_report_type=target_report_type,
            progress_callback=progress_callback
        )



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

