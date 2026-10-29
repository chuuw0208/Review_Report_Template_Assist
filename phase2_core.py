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
import io
import time
import shutil
import argparse
import copy
import posixpath
import zipfile
import datetime
import xml.etree.ElementTree as ET
from xml.sax.saxutils import quoteattr

try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None

try:
    import win32api
except ImportError:
    win32api = None

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
# CLIPBOARD-SAFE RANGE TRANSFER (COM)
# ============================================================================

CLIPBOARD_SETTLE_MS = 200   # pause after Copy/Paste so heavy images/charts finish
CLIPBOARD_RETRIES   = 3     # retries on transient "clipboard busy" COM errors


def _settle_clipboard(ms=CLIPBOARD_SETTLE_MS):
    """
    DoEvents() equivalent: pump pending Windows messages, then sleep briefly
    so the clipboard can finish rendering large images / MS Graph objects
    before the next Copy() or Paste().
    """
    if pythoncom is not None:
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass
    if win32api is not None:
        win32api.Sleep(ms)
    else:
        time.sleep(ms / 1000.0)


def copy_paste_section(reference_doc, template_doc, src_start, src_end, dest_pos,
                       retries=CLIPBOARD_RETRIES):
    """
    Copy reference_doc[src_start:src_end] and paste it at template_doc[dest_pos].

    Range.Copy()/Range.Paste() is Word's native clipboard transfer, so it
    carries everything anchored in the range: text, tables, inline shapes
    (pictures), floating shapes, charts and MS Graph / OLE objects.

    The paste point is a collapsed range, so nothing in the template is
    overwritten. With TrackRevisions ON the paste is recorded as an insertion.

    Raises the last COM error if every attempt fails.
    """
    last_exc = None
    for attempt in range(1, retries + 1):
        try:
            src = reference_doc.Range(src_start, src_end)
            src.Copy()
            _settle_clipboard()          # let the clipboard render images

            dest = template_doc.Range(dest_pos, dest_pos)
            dest.Paste()
            _settle_clipboard()          # let Word finish placing shapes
            return
        except Exception as exc:         # e.g. "OpenClipboard failed" / RPC busy
            last_exc = exc
            print(f"  [WARN] Copy/Paste attempt {attempt}/{retries} failed: {exc}")
            _settle_clipboard(CLIPBOARD_SETTLE_MS * attempt * 2)
    raise last_exc


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
            # Copy the reference body (text + images + charts) and INSERT it at
            # the template body start (collapsed range = no deletion), which
            # preserves all existing template body content below.
            copy_paste_section(
                reference_doc, template_doc,
                ref_h["body_start"], ref_h["body_end"],
                tmpl_h["body_start"],
            )
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
    Robust Affirmation tailoring for Word COM:
      1. Prunes Section 2 ('Additional Analysis since prior review' and all its subsections)
         BEFORE Track Changes is engaged, transforming the working copy into a clean Affirmation template.
      2. For Auto-Numbered headings (Word Multilevel Lists), Word automatically, natively,
         and instantly recalculates Section 3 -> 2, 3.1 -> 2.1, 3.2 -> 2.2, 3.3 -> 2.3!
      3. For Typed-Numbered headings (where numbers are static text), scans and updates
         typed text numbers (3. -> 2., 3.1 -> 2.1, etc.).
    """
    was_tracking = template_doc.TrackRevisions
    template_doc.TrackRevisions = False

    headings = build_section_map(template_doc)
    deleted_sec2 = False

    # Step 1: Find Section 2 and compute full range up to next Heading 1
    sec2_idx = None
    for idx, h in enumerate(headings):
        if "additional analysis since prior review" in h["key"] or "additional analysis" in h["key"]:
            sec2_idx = idx
            break

    if sec2_idx is not None:
        h = headings[sec2_idx]
        start_pos = h["heading_start"]
        end_pos = template_doc.Content.End
        for next_h in headings[sec2_idx + 1:]:
            if next_h["level"] == 1:
                end_pos = next_h["heading_start"]
                break
            else:
                end_pos = next_h["body_end"]

        print(f"\n  [AFFIRMATION] Removing Section 2 ('{h['text']}' and all subsections) to adapt template for Affirmation...")
        del_range = template_doc.Range(start_pos, end_pos)
        del_range.Delete()
        deleted_sec2 = True

    # Update Word fields & list numbering
    try:
        template_doc.Fields.Update()
    except Exception:
        pass

    # Step 2: Fallback for Typed-Numbered headings (e.g. static '3.', '3.1', '3.1.1')
    para_count = template_doc.Paragraphs.Count
    for i in range(1, para_count + 1):
        try:
            para = template_doc.Paragraphs.Item(i)
            style_name = para.Style.NameLocal
        except Exception:
            continue

        if style_name in HEADING_STYLE_NAMES:
            text = para.Range.Text
            m = re.match(r'^(\s*)3(\.[\d\.]*)(.*)', text)
            if m:
                leading = m.group(1)
                num_tail = m.group(2)
                rest = m.group(3)
                old_num = f"{leading}3{num_tail}"
                new_num = f"{leading}2{num_tail}"
                try:
                    f = para.Range.Find
                    f.ClearFormatting()
                    f.Text = old_num
                    f.Replacement.ClearFormatting()
                    f.Replacement.Text = new_num
                    f.Forward = True
                    f.Wrap = 0  # wdFindStop
                    f.Execute(Replace=1)  # wdReplaceOne
                    print(f"  [TYPED RENUMBER] '{text.strip()}' -> '{new_num}{rest.strip()}'")
                except Exception as e:
                    print(f"  [WARN] Failed to renumber heading '{text.strip()}': {e}")
            else:
                try:
                    if para.Range.ListFormat.ListType != 0:
                        print(f"  [AUTO-NUMBERED] '{text.strip()}' is auto-numbered by Word (ListString: {para.Range.ListFormat.ListString}).")
                except Exception:
                    pass

    template_doc.TrackRevisions = was_tracking
    return deleted_sec2



# ============================================================================
# TABLE CELL TARGETING (COM)
# ============================================================================
#
# Every Word table cell's Range.Text ends with the END-OF-CELL marker, which
# COM returns as "\r\x07" (CR + BEL). Inside a cell, "\r" also separates
# paragraphs and "\x0b" is a manual line break (Shift+Enter). None of these
# are visible content, so they must be stripped before comparing strings.

WD_CHARACTER       = 1          # wdCharacter (unit for Range.MoveEnd)
WD_COLOR_AUTOMATIC = -16777216  # wdColorAutomatic


def clean_cell_text(raw_text):
    """
    Return the visible text of a Word table cell read via COM.

        clean_cell_text("Model Component\\r\\x07")        -> "Model Component"
        clean_cell_text("Line 1\\rLine 2\\x0bMore\\r\\x07") -> "Line 1 Line 2 More"
    """
    if not raw_text:
        return ""
    # 1. Drop the trailing end-of-cell marker (any mix of \r and \x07).
    text = raw_text.rstrip("\r\x07")
    # 2. Internal paragraph / line breaks / stray BELs -> spaces.
    for ch in ("\r", "\x0b", "\x07"):
        text = text.replace(ch, " ")
    # 3. Collapse whitespace.
    return " ".join(text.split())


def find_table_by_header(doc, table_identifier_text, occurrence=1):
    """
    Return (table_index, table) for the Nth (occurrence, 1-based) table in
    doc.Tables whose Cell(1, 1) contains table_identifier_text
    (case/whitespace-insensitive). Returns (None, None) when no table matches.
    """
    needle = " ".join(str(table_identifier_text).lower().split())
    seen = 0
    for t_idx in range(1, doc.Tables.Count + 1):
        table = doc.Tables.Item(t_idx)
        try:
            header = clean_cell_text(table.Cell(1, 1).Range.Text)
        except Exception:
            continue  # Cell(1,1) unavailable (unusual merged layout)
        if needle in header.lower():
            seen += 1
            if seen == occurrence:
                return t_idx, table
    return None, None


def _clear_inherited_instruction_format(rng):
    """
    Text inserted into a cell inherits the formatting of the character before
    it. If that was red italic instruction text, the new value would look like
    an instruction, so reset italic + colour on the inserted range only.
    """
    try:
        font = rng.Font
        if font.Italic in (True, -1) and is_red(font.Color, font.ColorIndex):
            font.Italic = False
            font.Color = WD_COLOR_AUTOMATIC
    except Exception:
        pass


def modify_table_cell(doc, table_identifier_text, target_row, target_col, new_text,
                      occurrence=1):
    """
    Replace the content of one cell in the table identified by its Cell(1,1).

    Args:
        doc:                   Word Document (COM).
        table_identifier_text: text that Cell(1,1) must contain, e.g. "Model Component".
        target_row/target_col: 1-based cell coordinates (Word's Cell(row, col)).
        new_text:              str, or callable(old_text) -> str for per-cell
                               actions (e.g. append, prefix, conditional fill).
        occurrence:            which matching table to edit (1 = first).

    With TrackRevisions ON, the change is recorded as a deletion of the old
    text + insertion of the new text. Returns True if the cell was found
    (changed or already up to date), False otherwise.
    """
    t_idx, table = find_table_by_header(doc, table_identifier_text, occurrence)
    if table is None:
        print(f"  [WARN] No table #{occurrence} whose Cell(1,1) contains "
              f"\"{table_identifier_text}\".")
        return False

    try:
        cell = table.Cell(target_row, target_col)
    except Exception as exc:
        print(f"  [WARN] Table #{t_idx} (\"{table_identifier_text}\") has no "
              f"Cell({target_row},{target_col}): {exc}")
        return False

    old_text = clean_cell_text(cell.Range.Text)
    value = new_text(old_text) if callable(new_text) else new_text
    value = "" if value is None else str(value)

    if old_text == clean_cell_text(value):
        print(f"  [CELL] Table #{t_idx} Cell({target_row},{target_col}) already "
              f"\"{old_text}\" — unchanged.")
        return True

    # Cell content WITHOUT the end-of-cell marker. Shrinking by one character
    # means the marker itself is never deleted, so the table structure is safe.
    content = cell.Range
    content.MoveEnd(WD_CHARACTER, -1)
    if content.End > content.Start:
        content.Delete()                       # tracked deletion of old text

    # Insert at the point just before the end-of-cell marker.
    insert_pos = table.Cell(target_row, target_col).Range.End - 1
    inserted = doc.Range(insert_pos, insert_pos)
    inserted.InsertAfter(value)                # tracked insertion; range expands
    _clear_inherited_instruction_format(inserted)

    print(f"  [CELL] Table #{t_idx} Cell({target_row},{target_col}): "
          f"\"{old_text}\" -> \"{value}\"")
    return True


def apply_table_cell_rules(doc, rules):
    """
    Apply a list of per-cell rules to a COM document. Each rule is a dict:
        {"table": "Model Component", "row": 2, "col": 3, "text": "Approved"}
    where "text" may also be callable(old_text) -> new_text, and an optional
    "occurrence" (default 1) selects which matching table to edit.
    Returns the number of rules that found their target cell.
    """
    applied = 0
    for rule in rules or []:
        if modify_table_cell(doc, rule["table"], rule["row"], rule["col"], rule["text"],
                             occurrence=rule.get("occurrence", 1)):
            applied += 1
    print(f"  [INFO] Applied {applied}/{len(rules or [])} table cell rule(s).")
    return applied


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

        report_progress(25, "Opening reference and working documents...")

        # ---- Open documents ----
        reference_doc = word.Documents.Open(reference_path, ReadOnly=True)
        print(f"[INFO] Opened reference (read-only): {os.path.basename(reference_path)}")

        # Open working copy at output_path. Original template_path is NEVER opened for writing.
        template_doc = word.Documents.Open(output_path)
        print(f"[INFO] Opened working copy:          {os.path.basename(output_path)}")

        # ---- Target Report Type tailoring (Affirmation normalization) ----
        if str(target_report_type).strip().lower() == "affirmation":
            report_progress(30, "Tailoring template for Affirmation (removing Section 2)...")
            remove_affirmation_excluded_sections(template_doc)

        # ---- Track Changes ON ----
        template_doc.TrackRevisions = True
        print("[INFO] Track Changes -> ON\n")

        # ---- Build section maps ----
        report_progress(35, "Mapping document sections...")
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
        if child.tag == _w_tag("sectPr"):
            continue  # document-level page setup, never part of a section body
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

def tailor_affirmation_openxml(body_elem):
    """
    Robust Affirmation tailoring for OpenXML:
      1. Prunes Section 2 ('Additional Analysis since prior review' and all its subsections)
         cleanly from the document body before migration.
      2. For Auto-Numbered headings (headings with numPr), Word automatically,
         natively, and instantly recalculates Section 3 -> 2, 3.1 -> 2.1, 3.2 -> 2.2!
      3. For Typed-Numbered headings (where numbers are literal text in <w:t>),
         scans and renumbers '3.' -> '2.', '3.1' -> '2.1', '3.2' -> '2.2'.
    """
    children = list(body_elem)
    sec2_start_idx = None
    sec2_end_idx = len(children)

    for idx, child in enumerate(children):
        if child.tag == _P_TAG:
            lvl, text = get_heading_info_openxml(child)
            if lvl == 1:
                norm = normalize_heading(text)
                if "additional analysis since prior review" in norm or "additional analysis" in norm:
                    sec2_start_idx = idx
                elif sec2_start_idx is not None and idx > sec2_start_idx:
                    sec2_end_idx = idx
                    break

    if sec2_start_idx is not None:
        print(f"  [ROBUST AFFIRMATION] Removing Section 2 elements {sec2_start_idx} to {sec2_end_idx} to adapt template for Affirmation...")
        for i in range(sec2_end_idx - 1, sec2_start_idx - 1, -1):
            body_elem.remove(children[i])

    # Inspect remaining headings for typed numbers and renumber if needed
    renum_count = 0
    for child in body_elem.findall(_P_TAG):
        lvl, text = get_heading_info_openxml(child)
        if lvl is not None and text:
            m = re.match(r'^(\s*)3(\.[\d\.]*)(.*)', text)
            if m:
                leading = m.group(1)
                num_tail = m.group(2)
                rest = m.group(3)
                old_num = f"{leading}3{num_tail}"
                new_num = f"{leading}2{num_tail}"

                for r in child.findall(_R_TAG):
                    for t in r.findall(_T_TAG):
                        if t.text and re.match(r'^(\s*)3(\.[\d\.]*)', t.text):
                            t.text = re.sub(r'^(\s*)3(\.[\d\.]*)', f"\\g<1>2\\g<2>", t.text)
                            renum_count += 1
                            print(f"  [TYPED RENUMBER] '{old_num}{rest.strip()}' -> '{new_num}{rest.strip()}'")
                            break
            else:
                pPr = child.find(_PPR_TAG)
                if pPr is not None and pPr.find(_w_tag("numPr")) is not None:
                    print(f"  [AUTO-NUMBERED] '{text.strip()}' will be automatically numbered by Word (native list cascade).")

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

# ============================================================================
# OPENXML PACKAGE MERGE — images, charts, OLE objects, hyperlinks
# ============================================================================
#
# In a .docx, an image inside a paragraph is only a pointer:
#     <a:blip r:embed="rId7"/>  ->  word/_rels/document.xml.rels  ->  media/image1.png
# Cloning the paragraph XML alone leaves "rId7" pointing at nothing (or at a
# *different* image in the template). _PackageMerger re-creates every
# relationship referenced by migrated XML, copies the target part (and that
# part's own relationships, e.g. chart -> embedded workbook) under a
# collision-free name, and registers the content type.

R_NS        = OPENXML_NS['r']
PKG_REL_NS  = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS       = "http://schemas.openxmlformats.org/package/2006/content-types"
_R_ATTR_PREFIX = f"{{{R_NS}}}"
_DOCPR_TAG  = f"{{{OPENXML_NS['wp']}}}docPr"
_SECTPR_TAG = _w_tag("sectPr")
_REL_TAG      = f"{{{PKG_REL_NS}}}Relationship"
_CT_DEFAULT   = f"{{{CT_NS}}}Default"
_CT_OVERRIDE  = f"{{{CT_NS}}}Override"
_MAIN_PART    = "word/document.xml"


def _rels_path_for(part):
    """'word/charts/chart1.xml' -> 'word/charts/_rels/chart1.xml.rels'"""
    folder, name = posixpath.split(part)
    return posixpath.join(folder, "_rels", name + ".rels")


def _resolve_target(source_part, target):
    """Resolve a relationship Target relative to its source part."""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def _xml_bytes(root, default_ns):
    """
    Serialize a flat package part (.rels / [Content_Types].xml) with its
    namespace as the default xmlns, as Word expects. ElementTree's own
    default_namespace option rejects their unqualified attributes (Id, Type).
    """
    def local(tag):
        return tag.split("}", 1)[-1]

    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
           f'<{local(root.tag)} xmlns="{default_ns}">']
    for child in root:
        attrs = "".join(f" {k}={quoteattr(v)}" for k, v in child.attrib.items())
        out.append(f"<{local(child.tag)}{attrs}/>")
    out.append(f"</{local(root.tag)}>")
    return "".join(out).encode("utf-8")


class _PackageMerger:
    """Copies reference-package parts needed by migrated XML into the template."""

    def __init__(self, dst_files, src_files):
        self.dst = dst_files
        self.src = src_files
        self.part_map = {}   # src part path -> dst part path
        self.rid_map = {}    # src document rId -> dst document rId
        self.copied_parts = 0

        self.src_doc_rels = {
            rel.get("Id"): rel
            for rel in self._load_xml(src_files, _rels_path_for(_MAIN_PART), "Relationships", PKG_REL_NS)
        }
        self.dst_doc_rels = self._load_xml(dst_files, _rels_path_for(_MAIN_PART), "Relationships", PKG_REL_NS)
        self.src_ct = self._load_xml(src_files, "[Content_Types].xml", "Types", CT_NS)
        self.dst_ct = self._load_xml(dst_files, "[Content_Types].xml", "Types", CT_NS)

        self._used_rids = {rel.get("Id") for rel in self.dst_doc_rels}
        self._next_docpr = 1 + max(
            [int(e.get("id", "0")) for e in self._iter_dst_docpr() if e.get("id", "").isdigit()] or [0]
        )

    # ---- loading ----

    @staticmethod
    def _load_xml(files, path, root_name, ns):
        data = files.get(path)
        if data:
            return ET.fromstring(data)
        return ET.Element(f"{{{ns}}}{root_name}")

    def _iter_dst_docpr(self):
        data = self.dst.get(_MAIN_PART)
        if not data:
            return []
        return ET.fromstring(data).iter(_DOCPR_TAG)

    # ---- public API ----

    def import_element(self, elem):
        """Re-link r:* attributes and de-duplicate drawing ids on a cloned element."""
        for node in elem.iter():
            for attr, val in list(node.attrib.items()):
                if attr.startswith(_R_ATTR_PREFIX):
                    new_rid = self._map_doc_rid(val)
                    if new_rid:
                        node.set(attr, new_rid)
            if node.tag == _DOCPR_TAG:          # each drawing needs a unique id
                node.set("id", str(self._next_docpr))
                self._next_docpr += 1

    def finalize(self):
        """Write merged document rels + content types back into dst_files."""
        self.dst[_rels_path_for(_MAIN_PART)] = _xml_bytes(self.dst_doc_rels, PKG_REL_NS)
        self.dst["[Content_Types].xml"] = _xml_bytes(self.dst_ct, CT_NS)

    # ---- relationships ----

    def _new_rid(self):
        n = len(self._used_rids) + 1
        while f"rId{n}" in self._used_rids:
            n += 1
        rid = f"rId{n}"
        self._used_rids.add(rid)
        return rid

    def _map_doc_rid(self, src_rid):
        if src_rid in self.rid_map:
            return self.rid_map[src_rid]
        rel = self.src_doc_rels.get(src_rid)
        if rel is None:
            return None

        mode = rel.get("TargetMode")
        if mode == "External":                  # hyperlinks, linked pictures
            new_target = rel.get("Target")
        else:
            dst_part = self._copy_part(_resolve_target(_MAIN_PART, rel.get("Target")))
            if dst_part is None:
                return None
            new_target = posixpath.relpath(dst_part, posixpath.dirname(_MAIN_PART))

        new_rid = self._new_rid()
        new_rel = ET.SubElement(self.dst_doc_rels, _REL_TAG)
        new_rel.set("Id", new_rid)
        new_rel.set("Type", rel.get("Type"))
        new_rel.set("Target", new_target)
        if mode:
            new_rel.set("TargetMode", mode)
        self.rid_map[src_rid] = new_rid
        return new_rid

    # ---- parts ----

    def _unique_part_name(self, src_part):
        if src_part not in self.dst:
            return src_part
        stem, ext = posixpath.splitext(src_part)
        n = 1
        while f"{stem}_ref{n}{ext}" in self.dst:
            n += 1
        return f"{stem}_ref{n}{ext}"

    def _copy_part(self, src_part):
        """Copy a part (+ its own rels, recursively). Returns the dst part path."""
        if src_part in self.part_map:
            return self.part_map[src_part]
        if src_part not in self.src:
            print(f"  [WARN] Reference part missing, cannot copy: {src_part}")
            return None

        dst_part = self._unique_part_name(src_part)
        self.part_map[src_part] = dst_part      # register first (cycle-safe)
        self.dst[dst_part] = self.src[src_part]
        self._register_content_type(src_part, dst_part)
        self.copied_parts += 1

        src_rels = _rels_path_for(src_part)
        if src_rels in self.src:                # e.g. chart -> embedded .xlsx
            rels_root = ET.fromstring(self.src[src_rels])
            for rel in rels_root:
                if rel.get("TargetMode") == "External":
                    continue
                child_dst = self._copy_part(_resolve_target(src_part, rel.get("Target")))
                if child_dst:
                    rel.set("Target", posixpath.relpath(child_dst, posixpath.dirname(dst_part)))
            self.dst[_rels_path_for(dst_part)] = _xml_bytes(rels_root, PKG_REL_NS)
        return dst_part

    def _register_content_type(self, src_part, dst_part):
        src_name, dst_name = "/" + src_part, "/" + dst_part
        for ov in self.src_ct.iter(_CT_OVERRIDE):
            if ov.get("PartName", "").lower() == src_name.lower():
                new_ov = ET.SubElement(self.dst_ct, _CT_OVERRIDE)
                new_ov.set("PartName", dst_name)
                new_ov.set("ContentType", ov.get("ContentType"))
                return

        ext = posixpath.splitext(dst_part)[1].lstrip(".").lower()
        if any(d.get("Extension", "").lower() == ext for d in self.dst_ct.iter(_CT_DEFAULT)):
            return
        for d in self.src_ct.iter(_CT_DEFAULT):
            if d.get("Extension", "").lower() == ext:
                new_d = ET.Element(_CT_DEFAULT)
                new_d.set("Extension", d.get("Extension"))
                new_d.set("ContentType", d.get("ContentType"))
                self.dst_ct.insert(0, new_d)    # Defaults conventionally first
                return
        print(f"  [WARN] No content type found for {dst_part}")


def _strip_section_breaks(elem):
    """Remove paragraph-level section breaks (they carry reference headers/footers)."""
    for pPr in elem.iter(_PPR_TAG):
        for sect in pPr.findall(_SECTPR_TAG):
            pPr.remove(sect)


def _collect_namespaces(xml_bytes):
    """Return {prefix: uri} for every namespace declared in an XML part."""
    ns = {}
    for _event, (prefix, uri) in ET.iterparse(io.BytesIO(xml_bytes), events=("start-ns",)):
        ns.setdefault(prefix, uri)
    return ns


def _register_namespaces(ns_map):
    """Keep Word's original prefixes (w14, wp14, a, pic, ...) when re-serializing."""
    for prefix, uri in ns_map.items():
        if not prefix:
            continue
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:
            pass  # reserved prefix such as ns0


def _ensure_root_namespaces(xml_bytes, *ns_maps):
    """
    ElementTree drops namespace declarations that no element uses, but Word's
    mc:Ignorable="w14 wp14 ..." still references those prefixes; a missing
    declaration makes Word report the file as corrupt. Re-add them on the root.
    """
    text = xml_bytes.decode("utf-8")
    m = re.search(r"<(?![?!])([^\s>/]+)([^>]*)>", text)
    if not m:
        return xml_bytes
    declared = set(re.findall(r"xmlns:([\w.\-]+)=", m.group(2)))
    extra = []
    for ns_map in ns_maps:
        for prefix, uri in ns_map.items():
            if prefix and prefix not in declared:
                extra.append(f' xmlns:{prefix}="{uri}"')
                declared.add(prefix)
    if not extra:
        return xml_bytes
    return (text[:m.end(1)] + "".join(extra) + text[m.end(1):]).encode("utf-8")


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

    # Register the documents' own prefixes so ET re-serializes w14/wp14/a/pic
    # etc. unchanged (Word rejects files whose mc:Ignorable prefixes vanish).
    tmpl_ns = _collect_namespaces(tmpl_files["word/document.xml"])
    ref_ns = _collect_namespaces(ref_files["word/document.xml"])
    _register_namespaces(ref_ns)
    _register_namespaces(tmpl_ns)

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

    # ---- Target Report Type tailoring (Affirmation template adaptation) ----
    if str(target_report_type).strip().lower() == "affirmation":
        report_progress(30, "Tailoring template for Affirmation (removing Section 2)...")
        tailor_affirmation_openxml(tmpl_body)

    report_progress(35, "Parsing section structures...")

    ref_tree = ET.ElementTree(ET.fromstring(ref_files["word/document.xml"]))
    ref_root = ref_tree.getroot()
    ref_body = ref_root.find(_BODY_TAG)

    tmpl_sections = parse_docx_sections_openxml(tmpl_body)
    ref_sections = parse_docx_sections_openxml(ref_body)
    ref_sec_by_key = {s['key']: s for s in ref_sections}

    report_progress(50, "Migrating matched sections with Track Changes...")

    # Re-links images / charts / OLE / hyperlinks referenced by migrated XML
    merger = _PackageMerger(tmpl_files, ref_files)

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
                    _strip_section_breaks(clone)
                    merger.import_element(clone)
                    wrap_in_tracked_ins_openxml(clone, rev_id_gen)
                    tmpl_body.insert(insert_pos, clone)
                    insert_pos += 1
                migrated_count += 1
                print(f"  [MIGRATE] Migrated: {s['text']}")

    merger.finalize()
    if merger.copied_parts:
        print(f"  [INFO] Copied {merger.copied_parts} embedded part(s) "
              f"(images/charts/objects) from reference.")

    report_progress(75, "Scanning and deleting red-italic instructions...")
    del_inst_count = delete_red_italic_openxml(tmpl_body, rev_id_gen)
    print(f"  [INFO] Deleted {del_inst_count} red-italic instruction block(s).")

    report_progress(92, "Enabling Track Changes in settings...")


    settings_data = tmpl_files.get("word/settings.xml", None)
    if settings_data is not None:
        settings_ns = _collect_namespaces(settings_data)
        _register_namespaces(settings_ns)
        s_tree = ET.ElementTree(ET.fromstring(settings_data))
        s_root = s_tree.getroot()
        if s_root.find(_TRACK_REV_TAG) is None:
            ET.SubElement(s_root, _TRACK_REV_TAG)
        tmpl_files["word/settings.xml"] = _ensure_root_namespaces(
            ET.tostring(s_root, encoding="utf-8", xml_declaration=True), settings_ns)
    else:
        new_settings = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="{W_NS}"><w:trackRevisions/></w:settings>'
        tmpl_files["word/settings.xml"] = new_settings.encode("utf-8")

    tmpl_files["word/document.xml"] = _ensure_root_namespaces(
        ET.tostring(tmpl_root, encoding="utf-8", xml_declaration=True), tmpl_ns, ref_ns)

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

