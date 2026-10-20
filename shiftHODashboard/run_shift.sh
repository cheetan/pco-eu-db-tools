#!/bin/bash
# Schedules the next wake then runs the dashboard refresh.
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON=/opt/homebrew/bin/python3

bash "$SCRIPT_DIR/schedule_next_wake.sh"

caffeinate -i "$PYTHON" "$SCRIPT_DIR/dashboard.py" --refresh
