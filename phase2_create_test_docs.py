"""
Phase 2 – Test Fixture Generator.

Creates two Word documents for testing the Phase 2 migration pipeline:

  1. test_template_assessment.docx
     - Standard Assessment template with Heading 1/2 sections
     - Red italic instruction text in several sections
     - A placeholder table under "Review Type Determination"
     - Normal placeholder body text

  2. test_reference_assessment.docx
     - A "finalized" Assessment report with matching headings
     - Real content paragraphs replacing the placeholders
     - A filled-in table under "Review Type Determination"
     - Filename contains "Assessment" for auto-detection

Usage:
    python phase2_create_test_docs.py
"""

import os
import sys
import time
import win32com.client
import pythoncom

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Word constants
WD_FORMAT_XML_DOC = 12
WD_DO_NOT_SAVE    = 0
WD_COLLAPSE_END   = 0
WD_COLOR_RED      = 255   # RGB(255,0,0) in Word BGR format


# ============================================================================
# HELPER — DocBuilder
# ============================================================================

class DocBuilder:
    """Convenience wrapper for building a Word document paragraph-by-paragraph."""

    def __init__(self, doc):
        self.doc = doc
        self._used_first = False

    # ---- paragraph helpers ----

    def h1(self, text):
        self._add(text, "Heading 1")

    def h2(self, text):
        self._add(text, "Heading 2")

    def h3(self, text):
        self._add(text, "Heading 3")

    def body(self, text):
        self._add(text, "Normal")

    def instruction(self, text):
        """Red italic instruction text (the kind Phase 2 should delete)."""
        self._add(text, "Normal", italic=True, color=WD_COLOR_RED)

    def bold_body(self, text):
        self._add(text, "Normal", bold=True)

    # ---- table helper ----

    def table(self, data):
        """
        Insert a bordered table. `data` is a list of rows, each row a
        list of cell strings.  First row is bolded as a header.
        """
        doc = self.doc
        rows = len(data)
        cols = len(data[0]) if data else 0

        # Add an empty paragraph at the end for the table to occupy
        doc.Content.InsertParagraphAfter()
        rng = doc.Paragraphs(doc.Paragraphs.Count).Range

        tbl = doc.Tables.Add(rng, rows, cols)
        tbl.Borders.Enable = True

        for r, row_data in enumerate(data, 1):
            for c, cell_text in enumerate(row_data, 1):
                tbl.Cell(r, c).Range.Text = str(cell_text)

        # Bold header row
        tbl.Rows(1).Range.Font.Bold = True
        self._used_first = True  # we have content now

    # ---- internal ----

    def _add(self, text, style, bold=False, italic=False, color=None):
        doc = self.doc

        if not self._used_first:
            # Reuse the empty first paragraph that every new doc starts with
            para = doc.Paragraphs(1)
            self._used_first = True
        else:
            # Append a new paragraph at the end of the document
            doc.Content.InsertParagraphAfter()
            para = doc.Paragraphs(doc.Paragraphs.Count)

        # Insert text before the paragraph mark (preserves ¶)
        if text:
            para.Range.InsertBefore(text)

        # Apply paragraph style
        para.Style = style

        # Apply character-level formatting overrides
        if bold:
            para.Range.Font.Bold = True
        if italic:
            para.Range.Font.Italic = True
        if color is not None:
            para.Range.Font.Color = color


# ============================================================================
# DOCUMENT CREATORS
# ============================================================================

def create_test_template(word, output_path):
    """Build a test Assessment template with headings, red-italic text, and a table."""
    doc = word.Documents.Add()
    b = DocBuilder(doc)

    # ---- Title page (no heading style) ----
    b.bold_body("Model Risk Management")
    b.bold_body("Assessment Report Template")
    b.body("")

    # ---- Section 1 ----
    b.h1("1. Review Details")
    b.instruction(
        "Complete this section with the review details. "
        "Refer to the MRM Policy and Guidelines for instructions."
    )

    b.h2("1.1 Review Type Determination")
    b.instruction(
        "Select the appropriate review type from the table below "
        "and provide justification for the selected type."
    )
    b.table([
        ["Review Type",          "Description",                       "Selected"],
        ["Full Assessment",      "Complete model review",             ""],
        ["Targeted Assessment",  "Focused review of specific areas",  ""],
    ])
    b.body("Review type has not been determined yet.")

    b.h2("1.2 Review Scope")
    b.instruction(
        "Describe the scope of this review in detail. Include model "
        "components, data sources, and the testing approach."
    )
    b.body("Scope has not been defined yet.")

    b.h2("1.3 Review History")
    b.instruction(
        "Summarize the review history including prior findings and remediation status."
    )
    b.body("No prior review history has been documented.")

    # ---- Section 2 ----
    b.h1("2. Additional Analysis since prior review")
    b.instruction(
        "This section documents analysis performed since the prior review cycle."
    )

    b.h2("2.1 Key Challenges")
    b.instruction(
        "Document key challenges identified during the review period. "
        "Include data issues, methodology concerns, and implementation problems."
    )
    b.body("Key challenges have not been documented yet.")

    b.h2("2.2 Model Inputs, Assumptions & Limitations Review")
    b.instruction(
        "Provide a detailed review of model inputs, key assumptions, and known limitations."
    )
    b.body("Inputs and assumptions review pending.")

    # ---- Section 3 ----
    b.h1("3. Governance & Controls")

    b.h2("3.1 Model Controls")
    b.instruction(
        "Describe the model controls currently in place. "
        "Include both automated and manual controls."
    )
    b.body("Model controls have not been described yet.")

    b.h2("3.2 Performance Monitoring & Reporting")
    b.instruction(
        "Detail the performance monitoring framework and reporting cadence."
    )
    b.body("Performance monitoring details pending.")

    # ---- Save & close ----
    doc.SaveAs2(os.path.abspath(output_path), FileFormat=WD_FORMAT_XML_DOC)
    doc.Close(SaveChanges=WD_DO_NOT_SAVE)
    print(f"[OK] Template created: {output_path}")


def create_test_reference(word, output_path):
    """Build a test finalized Assessment reference report with real content."""
    doc = word.Documents.Add()
    b = DocBuilder(doc)

    # ---- Title page ----
    b.bold_body("Model XYZ — Assessment Report — Q3 2025")
    b.body("Model ID: MDL-2024-0042  |  Version 3.1")
    b.body("")

    # ---- Section 1 ----
    b.h1("1. Review Details")
    b.body(
        "This assessment covers Model XYZ (ID: MDL-2024-0042), version 3.1, "
        "used for commercial real-estate portfolio credit loss forecasting."
    )

    b.h2("1.1 Review Type Determination")
    b.body(
        "A full model assessment was determined based on the materiality of "
        "the model and the number of significant changes since the last review."
    )
    b.table([
        ["Review Type",          "Description",                       "Selected"],
        ["Full Assessment",      "Complete model review",             "Y"],
        ["Targeted Assessment",  "Focused review of specific areas",  "N"],
    ])
    b.body(
        "Justification: The model underwent a major recalibration in Q2 2025 "
        "affecting all PD and LGD segments, warranting a full assessment."
    )

    b.h2("1.2 Review Scope")
    b.body(
        "The scope of this review encompasses model conceptual soundness, "
        "data quality, implementation verification, and performance testing "
        "across all portfolio segments."
    )
    b.body(
        "Key areas of focus include the revised PD calibration methodology "
        "and the updated LGD cure-rate assumptions."
    )

    b.h2("1.3 Review History")
    b.body(
        "The model was previously assessed in Q1 2024 (Report ID: RPT-2024-011). "
        "Two medium-severity findings were identified: (F1) insufficient "
        "back-testing documentation and (F2) stale input data for the LGD module."
    )
    b.body(
        "Both findings were remediated by the model owner in Q3 2024 and "
        "verified by the MRM team."
    )

    # ---- Section 2 ----
    b.h1("2. Additional Analysis since prior review")
    b.body(
        "The following additional analysis has been performed since the "
        "Q1 2024 assessment."
    )

    b.h2("2.1 Key Challenges")
    b.body(
        "The primary challenge identified was data quality in the "
        "calibration sample for the commercial real-estate segment. "
        "Missing collateral values affected 12% of the sample, requiring "
        "imputation before recalibration."
    )

    b.h2("2.2 Model Inputs, Assumptions & Limitations Review")
    b.body(
        "Key inputs reviewed include macroeconomic scenario variables "
        "(GDP growth, unemployment rate, CRE price index). "
        "The assumption of log-normal LGD distribution was validated "
        "using a Kolmogorov–Smirnov test (p = 0.34)."
    )
    b.body(
        "Limitation: The model does not capture sector-specific "
        "concentration risk within the CRE portfolio."
    )

    # ---- Section 3 ----
    b.h1("3. Governance & Controls")

    b.h2("3.1 Model Controls")
    b.body(
        "Controls include quarterly back-testing of PD and LGD predictions, "
        "automated data-quality checks at the ETL stage, and manual "
        "override logging with dual-approval workflow."
    )

    b.h2("3.2 Performance Monitoring & Reporting")
    b.body(
        "Model performance is monitored through monthly KRI dashboards "
        "and quarterly MRM reports to the Model Risk Committee. "
        "Breach thresholds are set at ±15% deviation from expected loss."
    )

    # ---- Save & close ----
    doc.SaveAs2(os.path.abspath(output_path), FileFormat=WD_FORMAT_XML_DOC)
    doc.Close(SaveChanges=WD_DO_NOT_SAVE)
    print(f"[OK] Reference created: {output_path}")


# ============================================================================
# MAIN
# ============================================================================

def main():
    template_path  = os.path.join(SCRIPT_DIR, "test_template_assessment.docx")
    reference_path = os.path.join(SCRIPT_DIR, "test_reference_assessment.docx")

    word = None
    try:
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        print("[INFO] Word started.")

        create_test_template(word, template_path)
        create_test_reference(word, reference_path)

        print("\n[DONE] Test documents ready.  Run phase2_run_test.py next.\n")

    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)

    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        word = None
        pythoncom.CoUninitialize()
        time.sleep(1)
        print("[INFO] Word closed.")


if __name__ == "__main__":
    main()

