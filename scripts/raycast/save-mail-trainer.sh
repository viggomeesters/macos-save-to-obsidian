#!/bin/bash

# Required parameters:
# @raycast.schemaVersion 1
# @raycast.title Save Mail Trainer
# @raycast.mode compact
# @raycast.packageName Brain

# Optional parameters:
# @raycast.icon 🧭

# Documentation:
# @raycast.description Train Save Mail project detection rules
# @raycast.author Viggo

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
source "$SCRIPT_DIR/raycast-log.sh"

CMD="cd '$BRAIN_DIR' && python3 '$BRAIN_SCRIPTS/save_mail_trainer.py'"
osascript <<APPLESCRIPT
tell application "Terminal"
    activate
    do script "$CMD"
end tell
APPLESCRIPT

raycast_log "save-mail-trainer" "ok" "opened trainer"
echo "Save Mail Trainer geopend in Terminal"
