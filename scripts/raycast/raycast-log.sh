#!/bin/bash
# Shared logging for Raycast Script Commands.
# Source this, then call: raycast_log "command" "status" "summary"
#
# Delegates to _log_brain_event from brain-env.sh.

# Compute local repo root — scripts/raycast/ is 2 levels deep
_RAYCAST_LOG_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
export MACOS_MAIL_REPO="${_RAYCAST_LOG_DIR%/scripts/raycast}"
export BRAIN_DIR="$MACOS_MAIL_REPO"
export BRAIN_SCRIPTS="$MACOS_MAIL_REPO/scripts"
export BRAIN_SHARED="$MACOS_MAIL_REPO/shared"
export LIFE_OS_SHARED="$BRAIN_SHARED"
export RAYCAST_COMMAND_SCRIPT="$(cd "$(dirname "${BASH_SOURCE[1]:-$0}")" && pwd)/$(basename "${BASH_SOURCE[1]:-$0}")"
export RAYCAST_COMMAND_LOG="${BRAIN_STATE_DIR:-$BRAIN_DIR/state}/raycast/commands.jsonl"
export RAYCAST_PYTHON="${RAYCAST_PYTHON:-/usr/bin/python3}"

if [[ ! -d "$BRAIN_DIR" || ! -d "$BRAIN_SCRIPTS" || ! -d "$BRAIN_SHARED" ]]; then
    echo "ERROR: BRAIN_DIR resolution failed: $BRAIN_DIR" >&2
    exit 1
fi

_PYTHON_VERSION="$("$RAYCAST_PYTHON" -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")' 2>/dev/null || true)"
_PYTHON_USER_SITE="$("$RAYCAST_PYTHON" -c 'import site; print(site.getusersitepackages())' 2>/dev/null || true)"
_REPO_USER_HOME="${MACOS_MAIL_REPO%%/Dev/*}"
for _PY_SITE in \
    "$_PYTHON_USER_SITE" \
    "$_REPO_USER_HOME/Library/Python/$_PYTHON_VERSION/lib/python/site-packages"; do
    if [[ -n "$_PY_SITE" && -d "$_PY_SITE" ]]; then
        case ":${PYTHONPATH:-}:" in
            *":$_PY_SITE:"*) ;;
            *) export PYTHONPATH="$_PY_SITE${PYTHONPATH:+:$PYTHONPATH}" ;;
        esac
    fi
done
unset _PY_SITE _PYTHON_VERSION _PYTHON_USER_SITE _REPO_USER_HOME

# Ensure brain-env.sh is loaded (provides _log_brain_event)
if ! type _log_brain_event &>/dev/null; then
    if [[ -f "$BRAIN_SCRIPTS/brain-env.sh" ]]; then
        source "$BRAIN_SCRIPTS/brain-env.sh"
    fi
fi

if ! type _log_brain_event &>/dev/null; then
    _log_brain_event() { :; }
fi

_raycast_json_escape() {
    local value="${1:-}"
    value="${value//\\/\\\\}"
    value="${value//\"/\\\"}"
    value="${value//$'\n'/\\n}"
    value="${value//$'\r'/\\r}"
    printf '%s' "$value"
}

_raycast_command_line() {
    local rendered=""
    local arg
    for arg in "$@"; do
        if [[ -z "$rendered" ]]; then
            rendered="$(printf '%q' "$arg")"
        else
            rendered="$rendered $(printf '%q' "$arg")"
        fi
    done
    printf '%s' "$rendered"
}

_raycast_abs_path() {
    local path="${1:-}"
    if [[ -z "$path" ]]; then
        return
    fi
    "$RAYCAST_PYTHON" -c 'import os, sys; print(os.path.abspath(sys.argv[1]))' "$path" 2>/dev/null || printf '%s' "$path"
}

_raycast_resolve_target() {
    local executable="${1:-}"
    local first_arg="${2:-}"
    if [[ -z "$executable" ]]; then
        return
    fi
    case "$(basename "$executable")" in
        python|python3|bash|sh|zsh)
            if [[ -n "$first_arg" && "$first_arg" != -* && "$first_arg" == */* ]]; then
                _raycast_abs_path "$first_arg"
                return
            fi
            ;;
    esac
    if [[ "$executable" == */* ]]; then
        _raycast_abs_path "$executable"
    else
        command -v "$executable" 2>/dev/null || printf '%s' "$executable"
    fi
}

_raycast_now_epoch() {
    "$RAYCAST_PYTHON" -c 'import time; print(f"{time.time():.6f}")'
}

_raycast_now_utc() {
    "$RAYCAST_PYTHON" -c 'from datetime import datetime, timezone; print(datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z"))'
}

_raycast_now_local() {
    date "+%Y-%m-%d %H:%M:%S %Z"
}

_raycast_duration() {
    "$RAYCAST_PYTHON" - "$1" "$2" <<'PY'
import sys

start = float(sys.argv[1])
end = float(sys.argv[2])
print(f"{end - start:.2f}")
PY
}

_raycast_write_command_log() {
    local cmd="${1:-unknown}"
    local status="${2:-ok}"
    local summary="${3:-}"
    local target="${4:-}"
    local exit_code="${5:-}"
    local started_at="${6:-}"
    local ended_at="${7:-}"
    local duration_seconds="${8:-}"
    local ts
    ts=$(date -u +"%Y-%m-%dT%H:%M:%S.000Z")
    mkdir -p "$(dirname "$RAYCAST_COMMAND_LOG")" 2>/dev/null || true
    printf '{"timestamp":"%s","command":"%s","status":"%s","summary":"%s","script_path":"%s","target_path":"%s","cwd":"%s","exit_code":"%s","started_at":"%s","ended_at":"%s","duration_seconds":"%s"}\n' \
        "$ts" \
        "$(_raycast_json_escape "$cmd")" \
        "$(_raycast_json_escape "$status")" \
        "$(_raycast_json_escape "$summary")" \
        "$(_raycast_json_escape "$RAYCAST_COMMAND_SCRIPT")" \
        "$(_raycast_json_escape "$target")" \
        "$(_raycast_json_escape "$PWD")" \
        "$(_raycast_json_escape "$exit_code")" \
        "$(_raycast_json_escape "$started_at")" \
        "$(_raycast_json_escape "$ended_at")" \
        "$(_raycast_json_escape "$duration_seconds")" >> "$RAYCAST_COMMAND_LOG" 2>/dev/null || true
}

raycast_log() {
    local cmd="${1:-unknown}"
    local status="${2:-ok}"
    local summary="${3:-}"
    local target="${4:-}"
    local exit_code="${5:-}"
    local started_at="${6:-}"
    local ended_at="${7:-}"
    local duration_seconds="${8:-}"
    _raycast_write_command_log "$cmd" "$status" "$summary" "$target" "$exit_code" "$started_at" "$ended_at" "$duration_seconds"
    BRAIN_EVENT_AGENT="raycast" _log_brain_event "raycast_command" "$status" \
        "cmd=$cmd" \
        "summary=$summary" \
        "script_path=$RAYCAST_COMMAND_SCRIPT" \
        "target_path=$target" \
        "cwd=$PWD" \
        "command_log=$RAYCAST_COMMAND_LOG" \
        "exit_code=$exit_code" \
        "started_at=$started_at" \
        "ended_at=$ended_at" \
        "duration_seconds=$duration_seconds"
}

# Wrapper: run a command, capture output, log result, print output.
# Usage: raycast_run "command-name" python3 script.py --args
raycast_run() {
    local cmd_name="$1"; shift
    local target_path command_line started_at started_local start_epoch
    target_path="$(_raycast_resolve_target "$@")"
    command_line="$(_raycast_command_line "$@")"
    started_at="$(_raycast_now_utc)"
    started_local="$(_raycast_now_local)"
    start_epoch="$(_raycast_now_epoch)"
    echo "Running: $RAYCAST_COMMAND_SCRIPT"
    echo "Target:  $target_path"
    echo "Command: $command_line"
    echo "Log:     $RAYCAST_COMMAND_LOG"
    echo "Start:   $started_local"
    echo ""

    local output
    output=$("$@" 2>&1)
    local exit_code=$?
    local ended_at ended_local end_epoch duration_seconds
    ended_at="$(_raycast_now_utc)"
    ended_local="$(_raycast_now_local)"
    end_epoch="$(_raycast_now_epoch)"
    duration_seconds="$(_raycast_duration "$start_epoch" "$end_epoch")"

    if [ $exit_code -eq 0 ]; then
        local first_line
        first_line=$(echo "$output" | head -1)
        raycast_log "$cmd_name" "ok" "$first_line" "$target_path" "$exit_code" "$started_at" "$ended_at" "$duration_seconds"
    else
        raycast_log "$cmd_name" "error" "exit $exit_code" "$target_path" "$exit_code" "$started_at" "$ended_at" "$duration_seconds"
    fi

    echo "$output"
    echo ""
    echo "End:     $ended_local"
    echo "Elapsed: ${duration_seconds}s"
    return $exit_code
}

# Wrapper: stream output live, log result, and keep command metadata visible.
# Use this for long-running fullOutput commands where Raycast should show progress.
raycast_run_stream() {
    local cmd_name="$1"; shift
    local target_path command_line started_at started_local start_epoch
    target_path="$(_raycast_resolve_target "$@")"
    command_line="$(_raycast_command_line "$@")"
    started_at="$(_raycast_now_utc)"
    started_local="$(_raycast_now_local)"
    start_epoch="$(_raycast_now_epoch)"
    echo "Running: $RAYCAST_COMMAND_SCRIPT"
    echo "Target:  $target_path"
    echo "Command: $command_line"
    echo "Log:     $RAYCAST_COMMAND_LOG"
    echo "Start:   $started_local"
    echo ""

    "$@" 2>&1
    local exit_code=$?
    local ended_at ended_local end_epoch duration_seconds
    ended_at="$(_raycast_now_utc)"
    ended_local="$(_raycast_now_local)"
    end_epoch="$(_raycast_now_epoch)"
    duration_seconds="$(_raycast_duration "$start_epoch" "$end_epoch")"

    if [ $exit_code -eq 0 ]; then
        raycast_log "$cmd_name" "ok" "complete" "$target_path" "$exit_code" "$started_at" "$ended_at" "$duration_seconds"
    else
        raycast_log "$cmd_name" "error" "exit $exit_code" "$target_path" "$exit_code" "$started_at" "$ended_at" "$duration_seconds"
    fi

    echo ""
    echo "End:     $ended_local"
    echo "Elapsed: ${duration_seconds}s"
    return $exit_code
}
