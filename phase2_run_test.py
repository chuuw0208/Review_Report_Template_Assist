"""
Phase 2 – Test Runner.

Convenience script: runs the migration pipeline on the generated
test documents (test_template_assessment.docx + test_reference_assessment.docx)
and produces test_output_assessment.docx.

Usage:
    python phase2_run_test.py

Pre-requisite:
    Run  phase2_create_test_docs.py  first to generate the test files.
"""

import os
import sys

# Ensure the script directory is on the path so we can import phase2_core
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from phase2_core import run_migration

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

run_migration(
    template_path  = os.path.join(SCRIPT_DIR, "test_template_assessment.docx"),
    reference_path = os.path.join(SCRIPT_DIR, "test_reference_assessment.docx"),
    output_path    = os.path.join(SCRIPT_DIR, "test_output_assessment.docx"),
)

