#!/bin/bash
# ==============================================================================
# MRMV Report Content Migration Tool - macOS Launcher
# Double-click this file on macOS to launch the application!
# ==============================================================================

DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$DIR"

# Check if python3 is available
if command -v python3 >/dev/null 2>&1; then
    python3 "$DIR/gui_app.py"
elif command -v python >/dev/null 2>&1; then
    python "$DIR/gui_app.py"
else
    osascript -e 'display alert "Python 3 Not Found" message "Please install Python 3 on your Mac to run this tool."'
    exit 1
fi

