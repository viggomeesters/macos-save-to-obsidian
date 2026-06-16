#!/bin/bash

# Required parameters:
# @raycast.schemaVersion 1
# @raycast.title Save Mail
# @raycast.mode fullOutput
# @raycast.packageName Brain

# Optional parameters:
# @raycast.icon 📧

# Documentation:
# @raycast.description Save selected Apple Mail or Outlook message to vault; flagged mail creates a task
# @raycast.author Viggo

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
source "$SCRIPT_DIR/raycast-log.sh"
raycast_run_stream "save-mail" python3 "$BRAIN_SCRIPTS/save_mail.py"
