#!/bin/bash

# Required parameters:
# @raycast.schemaVersion 1
# @raycast.title Save Outlook Mail
# @raycast.mode fullOutput
# @raycast.packageName Brain

# Optional parameters:
# @raycast.icon 📬

# Documentation:
# @raycast.description Save selected Outlook message to vault; archives inbox mail after save
# @raycast.author Viggo

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
source "$SCRIPT_DIR/raycast-log.sh"
raycast_run_stream "save-outlook-mail" "$RAYCAST_PYTHON" "$BRAIN_SCRIPTS/save_mail.py" --client outlook
