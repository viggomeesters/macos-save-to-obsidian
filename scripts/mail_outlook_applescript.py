#!/usr/bin/env python3
"""Microsoft Outlook AppleScript wrappers for selected-mail save support."""

from __future__ import annotations

import hashlib
import html as html_lib
import re
import os
import site
import subprocess
import sys
import time
import uuid
from pathlib import Path
from typing import Any

from mail_applescript import _parse_headers, classify_mailbox_type

_LAST_NEW_OUTLOOK_RECORD: dict[str, Any] | None = None


def is_outlook_running() -> bool:
    """Check if Microsoft Outlook is running."""
    try:
        result = subprocess.run(
            ["pgrep", "-x", "Microsoft Outlook"], capture_output=True, timeout=5
        )
        return result.returncode == 0
    except Exception:
        return False


def run_applescript(script: str, timeout: int = 10) -> str:
    """Run AppleScript and return stdout."""
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired as e:
        raise TimeoutError(f"Outlook AppleScript timed out after {timeout}s") from e
    if result.returncode != 0:
        stderr = result.stderr.strip()
        if "No selected Outlook mail" in stderr:
            raise RuntimeError(
                "No selected Outlook mail. "
                "Check that Outlook is open, the BAM mailbox is visible, and one message is selected."
            )
        if "Select exactly one Outlook mail" in stderr:
            raise RuntimeError("Select exactly one Outlook mail.")
        if "not permitted" in stderr.lower() and "system events" in stderr.lower():
            raise RuntimeError(
                "System Events access is blocked. Allow Keyboard & Accessibility access for this app/script host."
            )
        if "-1743" in stderr or "not authorized" in stderr.lower():
            raise RuntimeError(
                "Microsoft Outlook Automation permission is missing for osascript"
            )
        raise RuntimeError(f"Outlook AppleScript error: {stderr}")
    return result.stdout.strip()


def _is_new_outlook_enabled() -> bool:
    """Avoid AppleScript activation against New Outlook; it uses a separate UI/runtime."""
    for key in ("RunningNewOutlook", "IsRunningNewOutlook", "EnableNewOutlook"):
        try:
            result = subprocess.run(
                ["defaults", "read", "com.microsoft.Outlook", key],
                capture_output=True,
                text=True,
                timeout=2,
            )
        except Exception:
            continue
        if result.returncode == 0 and result.stdout.strip() in {"1", "true", "YES"}:
            return True
    return False


def _clean_message_id(value: str) -> str:
    return (value or "").strip().strip("<>").strip()


def _split_outlook_list(value: str) -> list[str]:
    return [
        part.strip()
        for part in re.split(r",|;", value or "")
        if part and part.strip()
    ]


def _extract_email(value: str) -> str:
    match = re.search(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", value or "", re.I)
    return match.group(0).lower() if match else (value or "").strip().lower()


def _contains_email(value: str) -> bool:
    return bool(re.search(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", value or "", re.I))


def _stable_outlook_message_id(
    outlook_id: str,
    subject: str,
    sender_email: str,
    date_str: str,
) -> str:
    """Build a deterministic fallback id when Outlook does not expose Message-ID."""
    source = outlook_id.strip() or f"{subject}|{sender_email}|{date_str}"
    digest = hashlib.sha256(source.encode("utf-8")).hexdigest()[:24]
    return f"outlook-{digest}@local.outlook"


def _message_id_from_parts(
    outlook_id: str,
    all_headers: str,
    subject: str,
    sender_email: str,
    date_str: str,
) -> str:
    parsed = _parse_headers(all_headers)
    header_id = _clean_message_id(parsed.get("message-id", ""))
    if header_id:
        return header_id
    cleaned_outlook_id = _clean_message_id(outlook_id)
    if cleaned_outlook_id and "@" in cleaned_outlook_id and " " not in cleaned_outlook_id:
        return cleaned_outlook_id
    return _stable_outlook_message_id(
        cleaned_outlook_id, subject, sender_email, date_str
    )


def _parse_bool(value: str) -> bool:
    return (value or "").strip().lower() in {"true", "yes", "1"}


def _pick_from_header(headers: dict[str, str], keys: tuple[str, ...], default: str = "") -> str:
    for key in keys:
        value = headers.get(key)
        if value:
            return value
    return default


def _clean_address_list(value: str) -> str:
    return ", ".join(_extract_email(part) for part in _split_outlook_list(value) if _extract_email(part))


def _extract_display_from_header(value: str) -> str:
    if not value:
        return ""
    match = re.match(r'^\s*([^<"]+?)\s*<[^>]+>\s*$', value)
    if match:
        return match.group(1).strip().strip('"').strip("'")
    match = re.match(r'^"([^"]+)"\s*<[^>]+>$', value)
    if match:
        return match.group(1).strip()
    return value.strip()


def _clipboard_paste() -> str:
    """Return current clipboard content as text."""
    try:
        result = subprocess.run(["/usr/bin/pbpaste"], capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            return result.stdout
    except Exception:
        pass
    return ""


def _clipboard_paste_prefer_rich() -> str:
    """Return clipboard content from multiple formats, preferring parsed text."""
    candidates: list[str] = []
    commands = (
        ["/usr/bin/pbpaste"],
        ["/usr/bin/pbpaste", "-Prefer", "txt"],
        ["/usr/bin/pbpaste", "-Prefer", "rtf"],
    )
    for cmd in commands:
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=6)
            if result.returncode == 0:
                text = _clipboard_text_for_parsing(result.stdout or "").strip()
                if text:
                    candidates.append(text)
        except Exception:
            continue
    for candidate in candidates:
        if _looks_like_message_copy(candidate):
            return candidate
    return candidates[0] if candidates else _clipboard_text_for_parsing(_clipboard_paste())


def _set_clipboard_text(value: str) -> None:
    try:
        subprocess.run(
            ["/usr/bin/pbcopy"],
            input=value,
            text=True,
            check=False,
            timeout=5,
        )
    except Exception:
        pass


def _strip_rtf_markup(value: str) -> str:
    text = (value or "").strip()
    if not text.startswith("{\\rtf"):
        return value

    def decode_hex(match: re.Match[str]) -> str:
        try:
            return bytes([int(match.group(1), 16)]).decode("cp1252")
        except Exception:
            return ""

    def decode_unicode(match: re.Match[str]) -> str:
        try:
            codepoint = int(match.group(1))
            if codepoint < 0:
                codepoint += 65536
            return chr(codepoint)
        except Exception:
            return ""

    text = re.sub(r"\\u(-?\d+)\??", decode_unicode, text)
    text = re.sub(r"\\'([0-9a-fA-F]{2})", decode_hex, text)
    text = re.sub(r"\\(?:par|line)\b ?", "\n", text)
    text = re.sub(r"\\tab\b ?", "\t", text)
    text = re.sub(r"\\[{}\\]", lambda match: match.group(0)[1], text)
    text = re.sub(r"\\[a-zA-Z]+-?\d* ?", "", text)
    text = text.replace("{", "").replace("}", "")
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def _strip_html_markup(value: str) -> str:
    text = value or ""
    if not re.search(r"(?is)<(?:html|body|div|p|br|table|tr|td|th|span)\b", text):
        return value
    text = re.sub(r"(?is)<\s*br\s*/?\s*>", "\n", text)
    text = re.sub(r"(?is)</\s*(?:p|div|tr|table|h[1-6])\s*>", "\n", text)
    text = re.sub(r"(?is)</\s*(?:td|th)\s*>", "\t", text)
    text = re.sub(r"(?is)<[^>]+>", "", text)
    return html_lib.unescape(text).strip()


def _clipboard_text_for_parsing(raw_text: str) -> str:
    """Normalize pasteboard payloads to text before header detection."""
    text = raw_text or ""
    text = _strip_rtf_markup(text)
    text = _strip_html_markup(text)
    return text.replace("\r\n", "\n").replace("\r", "\n")


def _clipboard_payload_diagnostic(raw_text: str) -> str:
    normalized = _clipboard_text_for_parsing(raw_text).strip()
    raw = (raw_text or "").strip()
    kinds: list[str] = []
    if raw.startswith("{\\rtf"):
        kinds.append("rtf")
    if re.search(r"(?is)<(?:html|body|div|p|br|table|tr|td|th|span)\b", raw):
        kinds.append("html")
    if normalized.startswith("Running: ") or normalized.startswith("▶ Script:"):
        kinds.append("raycast-output")
    if normalized.startswith("__save-mail-verify-"):
        kinds.append("marker")
    if _contains_email(normalized):
        kinds.append("email-like")
    if re.search(
        r"(?im)^\s*(from|van|sent|verzonden|to|aan|subject|onderwerp|date|datum)\b",
        normalized,
    ):
        kinds.append("header-like")
    if not normalized:
        kinds.append("empty")
    return (
        f"clipboard chars={len(normalized)}, "
        f"lines={normalized.count(chr(10)) + 1 if normalized else 0}, "
        f"kind={'+'.join(kinds) if kinds else 'unknown'}"
    )


def _outlook_ui_actions_enabled() -> bool:
    return os.environ.get("SAVE_MAIL_OUTLOOK_UI_ACTIONS", "").strip().lower() in {
        "1",
        "true",
        "yes",
        "on",
    }


def _outlook_clipboard_fallback_enabled() -> bool:
    return os.environ.get("SAVE_MAIL_OUTLOOK_CLIPBOARD_FALLBACK", "").strip().lower() in {
        "1",
        "true",
        "yes",
        "on",
    }


def _looks_like_message_copy(raw_text: str) -> bool:
    text = _clipboard_text_for_parsing(raw_text).strip()
    if not text:
        return False
    if text.startswith("Running: ") or text.startswith("▶ Script:"):
        return False
    lowered = text.lower()
    if text.startswith("__save-mail-verify-"):
        return False
    header_markers = (
        "from:",
        "van:",
        "sent:",
        "to:",
        "aan:",
        "cc:",
        "subject:",
        "onderwerp:",
        "date:",
        "datum:",
        "verzonden:",
    )
    for marker in header_markers:
        if re.search(rf"(?m)^\s*{re.escape(marker)}", lowered):
            return True
    # Common New Outlook copy variants sometimes emit headers without the ":" delimiter.
    loose_markers = (
        r"^from\s+",
        r"^van\s+",
        r"^sent\s+",
        r"^to\s+",
        r"^aan\s+",
        r"^subject\s+",
        r"^onderwerp\s+",
        r"^date\s+",
        r"^datum\s+",
        r"^verzonden\s+",
    )
    for marker in loose_markers:
        if re.search(rf"(?m)^{marker}", lowered):
            return True
    # Fallback for copied raw payloads that keep full RFC headers in one block.
    if "message-id" in lowered and "@" in raw_text:
        return True
    # Very loose fallback: long payload + some email-like text can still be valid.
    if (
        len(text) > 200
        and _contains_email(text)
        and any(ch in text for ch in ("<", ">"))
    ):
        return True
    # Shorter body-only payloads with explicit date/from context can still be useful.
    if re.search(
        r"(?im)^(from|van|to|aan|subject|onderwerp|date|datum|verzonden)\s*:\s*.+",
        text,
    ):
        return True
    if (
        _contains_email(text)
        and len(text) > 80
        and (
            "@bam.com" in lowered
            or "@outlook." in lowered
            or "@microsoft" in lowered
        )
    ):
        return True
    return False


def _extract_freeform_header(raw_text: str, keys: tuple[str, ...]) -> str:
    for line in _clipboard_text_for_parsing(raw_text).splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        for key in keys:
            delimiter = r"(?:[:：]|\t| {2,}|\s+)" if len(key) > 3 else r"(?:[:：]|\t| {2,})"
            match = re.match(
                rf"(?i)^\s*{re.escape(key)}\s*{delimiter}\s*(.+)$",
                stripped,
            )
            if match:
                return match.group(1).strip()
    return ""


def _copy_selected_outlook_message_via_clipboard() -> str:
    """Copy selected message using system keystroke fallback and return clipboard text."""
    previous_clipboard = _clipboard_paste()
    previous_clipboard_clean = previous_clipboard.strip()
    verification_marker = f"__save-mail-verify-{uuid.uuid4()}"
    scripts = [
        r'''
tell application "Microsoft Outlook"
    activate
end tell

tell application "System Events"
    repeat 30 times
        try
            if frontmost of process "Microsoft Outlook" is true then exit repeat
        end try
        delay 0.1
    end repeat

    set outlookWasFrontmost to false
    try
        set outlookWasFrontmost to frontmost of process "Microsoft Outlook"
    end try
    if outlookWasFrontmost is false then error "Microsoft Outlook did not become frontmost before copy."

    tell process "Microsoft Outlook"
        set frontmost to true
        delay 0.1
        keystroke "c" using {command down}
    end tell
end tell
''',
    ]

    if _outlook_ui_actions_enabled():
        scripts.append(
            r'''
delay 0.15

tell application "System Events"
    tell process "Microsoft Outlook"
        tell menu bar 1
            try
                click menu bar item "Edit"
            on error
                click menu bar item "Bewerken"
            end try
        end tell
        delay 0.1
        try
            click (first menu item whose name is "Copy") of menu 1 of menu bar item "Edit" of menu bar 1
        on error
            try
                click (first menu item whose name is "Kopie") of menu 1 of menu bar item "Bewerken" of menu bar 1
            on error
                click (first menu item whose name is "Kopiëren") of menu 1 of menu bar item "Bewerken" of menu bar 1
            end try
        end try
    end tell
end tell
'''
        )
    last_error: str | None = None
    found_message_copy = False
    copied_text = ""
    marker_was_set = False
    try:
        _set_clipboard_text(verification_marker)
        marker_was_set = (
            _clipboard_text_for_parsing(_clipboard_paste()).strip()
            == verification_marker
        )
        for attempt in range(1, 6):
            for clipboard_script in scripts:
                try:
                    run_applescript(clipboard_script, timeout=12)
                    time.sleep(0.20 + (attempt * 0.15))
                    copied_text = _clipboard_paste_prefer_rich()
                    if copied_text:
                        stale_previous = (
                            not marker_was_set
                            and copied_text
                            in {previous_clipboard, previous_clipboard_clean}
                        )
                        if copied_text != verification_marker and not stale_previous:
                            if _looks_like_message_copy(copied_text):
                                found_message_copy = True
                                return copied_text
                            # Even when headers are fuzzy, keep last payload as a last resort.
                            last_error = ""
                        else:
                            last_error = "copied clipboard empty"
                    else:
                        last_error = "copied clipboard empty"
                except Exception as exc:
                    last_error = str(exc)
            time.sleep(0.15)
        if copied_text:
            # New Outlook can return slightly malformed but still useful payloads.
            stale_previous = (
                not marker_was_set
                and copied_text in {previous_clipboard, previous_clipboard_clean}
            )
            if copied_text != verification_marker and not stale_previous:
                if not (
                    copied_text.startswith("Running:")
                    or copied_text.startswith("▶ Script:")
                ):
                    found_message_copy = True
                    return copied_text
                last_error = "copied clipboard content not recognized as message"
            else:
                last_error = "No selected Outlook mail. Ensure one message is selected in New Outlook."
    finally:
        if previous_clipboard or previous_clipboard == "":
            _set_clipboard_text(previous_clipboard)
    if not found_message_copy:
        if last_error and "System Events access is blocked" in last_error:
            raise RuntimeError(last_error)
        if copied_text:
            raise RuntimeError(
                "Copied clipboard content is not an Outlook message "
                f"({_clipboard_payload_diagnostic(copied_text)}). "
                "Ensure one message is selected in New Outlook and retry."
            )
        raise RuntimeError("No selected Outlook mail. Ensure one message is selected in New Outlook.")
    return copied_text


def _parse_clipboard_headers(raw_text: str) -> tuple[dict[str, str], str, str]:
    normalized_text = _clipboard_text_for_parsing(raw_text)
    lines = normalized_text.split("\n")
    headers: dict[str, str] = {}
    raw_header_lines: list[str] = []
    body_lines: list[str] = []
    in_headers = False
    current_header = ""
    key_map = {
        "van": "from",
        "from": "from",
        "verzonden": "date",
        "datum": "date",
        "aan": "to",
        "to": "to",
        "cc": "cc",
        "bcc": "bcc",
        "onderwerp": "subject",
        "subject": "subject",
        "date": "date",
        "sent": "date",
        "date-and-time": "date",
        "message-id": "message-id",
        "in-reply-to": "in-reply-to",
        "references": "references",
        "thread-topic": "thread-topic",
        "content-type": "content-type",
        "mime-version": "mime-version",
    }

    for line in lines:
        if not in_headers:
            if not line.strip():
                continue
            in_headers = True
        if in_headers:
            if line.strip() == "":
                in_headers = False
                continue
            header_match = re.match(
                r"^\s*([A-Za-zÀ-ÿ0-9][A-Za-zÀ-ÿ0-9 -]{0,40})\s*"
                r"(?:[:：]|\t| {2,})\s*(.*)$",
                line,
            )
            if header_match:
                raw_key = header_match.group(1).strip().lower().replace(" ", "-")
                mapped_key = key_map.get(raw_key, raw_key)
                is_supported_header = (
                    mapped_key in key_map.values()
                    or raw_key.startswith("x-")
                )
                if is_supported_header:
                    current_header = mapped_key
                    headers[current_header] = header_match.group(2).strip()
                    raw_header_lines.append(
                        f"{current_header}: {headers[current_header]}"
                    )
                    continue
            if current_header and (line.startswith(" ") or line.startswith("\t")):
                headers[current_header] = f"{headers[current_header]} {line.strip()}"
                if raw_header_lines:
                    raw_header_lines[-1] = f"{current_header}: {headers[current_header]}"
                continue
            in_headers = False
            if line:
                body_lines.append(line)
        else:
            body_lines.append(line)

    return headers, "\n".join(raw_header_lines).strip(), "\n".join(body_lines).strip()


def _coerce_outlook_record_from_clipboard(raw_text: str) -> dict[str, Any] | None:
    if not _looks_like_message_copy(raw_text):
        return None

    headers, raw_headers, body = _parse_clipboard_headers(raw_text)
    if not headers and not body:
        return None

    subject = (
        _pick_from_header(headers, ("subject",), "")
        or _extract_freeform_header(raw_text, ("subject", "onderwerp"))
    )
    sender_raw = (
        _pick_from_header(headers, ("from",), "")
        or _extract_freeform_header(raw_text, ("from", "van"))
    )
    sender_display = _extract_display_from_header(sender_raw)
    sender_email = _extract_email(sender_raw)
    if not sender_email:
        sender_email = _extract_email(
            _pick_from_header(headers, ("from",), "")
            or _extract_freeform_header(raw_text, ("from", "van"))
        )

    account = "Microsoft Outlook"
    mailbox = _pick_from_header(headers, ("x-folder", "x-owa-folder", "folder"), "Outlook")
    date_str = (
        _pick_from_header(headers, ("date",), "")
        or _extract_freeform_header(raw_text, ("date", "datum", "verzonden"))
    )
    to_emails = _clean_address_list(
        _pick_from_header(headers, ("to",), "") or _extract_freeform_header(raw_text, ("to", "aan"))
    )
    cc_emails = _clean_address_list(
        _pick_from_header(headers, ("cc", "cc-list"), "")
        or _extract_freeform_header(raw_text, ("cc",))
    )
    all_headers = raw_headers or raw_text
    message_id = _message_id_from_parts("", all_headers, subject, sender_email, date_str)
    if not subject and not sender_email and not date_str:
        if len(body.strip()) < 120:
            return None
        if not any(
            key in headers
            for key in ("from", "to", "cc", "subject", "date")
        ):
            return None

    return {
        "subject": subject,
        "sender_email": sender_email,
        "sender_display": sender_display,
        "date_str": date_str,
        "message_id": message_id,
        "account": account,
        "mailbox": mailbox or "Outlook",
        "to": to_emails,
        "cc": cc_emails,
        "att_count": "0",
        "att_names": "",
        "is_flagged": False,
        "has_calendar": False,
        "all_headers": all_headers,
        "mailbox_type": classify_mailbox_type(mailbox, account),
        "body": body[:50000],
        "mail_client": "outlook",
    }


def _ax_attr(ax_module: Any, element: Any, name: str) -> Any:
    try:
        err, value = ax_module.AXUIElementCopyAttributeValue(element, name, None)
    except Exception:
        return None
    return value if err == 0 else None


def _ax_text(element: Any, ax_module: Any) -> str:
    values: list[str] = []
    for attr_name in ("AXDescription", "AXValue", "AXTitle"):
        value = _ax_attr(ax_module, element, attr_name)
        if isinstance(value, str) and value.strip() and value.strip() not in values:
            values.append(value.strip())
    return " ".join(values).strip()


def _ax_walk(
    ax_module: Any,
    element: Any,
    *,
    role: str,
    max_depth: int = 12,
    max_nodes: int = 5000,
) -> list[Any]:
    found: list[Any] = []
    visited = 0

    def visit(node: Any, depth: int) -> None:
        nonlocal visited
        visited += 1
        if depth > max_depth or visited > max_nodes:
            return
        if _ax_attr(ax_module, node, "AXRole") == role:
            found.append(node)
        children = _ax_attr(ax_module, node, "AXChildren") or []
        for child in list(children)[:300]:
            visit(child, depth + 1)

    visit(element, 0)
    return found


def _ax_row_text(ax_module: Any, row: Any) -> str:
    parts: list[str] = []
    for cell in list(_ax_attr(ax_module, row, "AXChildren") or []):
        text = _ax_text(cell, ax_module)
        if text:
            parts.append(text)
    if not parts:
        row_text = _ax_text(row, ax_module)
        if row_text:
            parts.append(row_text)
    return " ".join(parts).strip()


def _looks_like_new_outlook_message_row(row_text: str) -> bool:
    text = re.sub(r"\s+", " ", row_text or "").strip()
    if len(text) < 40:
        return False
    return bool(
        re.search(r"\b\d{1,2}/\d{1,2}/\d{4}\b", text)
        or re.search(
            r"\b\d{1,2}\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+\d{4}\b",
            text,
            re.I,
        )
    )


def _parse_new_outlook_ax_row(row_text: str, window_title: str = "") -> dict[str, Any] | None:
    text = re.sub(r"\s+", " ", row_text or "").strip()
    if not _looks_like_new_outlook_message_row(text):
        return None

    date_match = re.search(
        r"\b(\d{1,2}/\d{1,2}/\d{4}"
        r"(?:\s+\d{1,2}:\d{2})?"
        r"|\d{1,2}\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+\d{4}"
        r"(?:\s+\d{1,2}:\d{2})?)\b",
        text,
        re.I,
    )
    if not date_match:
        return None

    header_text = text[: date_match.start()].strip(" ,")
    date_str = date_match.group(1).strip()
    body = text[date_match.end() :].strip(" ,")
    header_parts = [part.strip() for part in header_text.split(",") if part.strip()]
    if header_parts and re.fullmatch(r"\d+\s+messages?", header_parts[0], re.I):
        header_parts = header_parts[1:]
    if len(header_parts) >= 2:
        subject = header_parts[-1]
        sender_display = ", ".join(header_parts[:-1])
    else:
        subject = header_text
        sender_display = ""

    if not subject:
        return None

    sender_email = _extract_email(sender_display)
    mailbox = "Outlook"
    account = "Microsoft Outlook"
    if window_title:
        title_parts = [part.strip() for part in window_title.split("•", 1)]
        if title_parts and title_parts[0]:
            mailbox = title_parts[0]
        if len(title_parts) > 1 and title_parts[1]:
            account = title_parts[1]

    all_headers = "\n".join(
        [
            f"From: {sender_display}",
            f"Subject: {subject}",
            f"Date: {date_str}",
            "X-Source: Microsoft Outlook Accessibility",
        ]
    )
    message_id = _message_id_from_parts("", all_headers, subject, sender_email, date_str)
    return {
        "subject": subject,
        "sender_email": sender_email,
        "sender_display": sender_display,
        "date_str": date_str,
        "message_id": message_id,
        "account": account,
        "mailbox": mailbox or "Outlook",
        "to": "",
        "cc": "",
        "att_count": "0",
        "att_names": "",
        "is_flagged": False,
        "has_calendar": False,
        "all_headers": all_headers,
        "mailbox_type": classify_mailbox_type(mailbox, account),
        "body": body[:50000],
        "mail_client": "outlook",
        "capture_source": "outlook-accessibility",
    }


def _get_selected_mail_header_from_accessibility() -> dict[str, Any]:
    checked_paths: list[str] = []
    try:
        import ApplicationServices as AX
    except Exception:
        version = f"{sys.version_info.major}.{sys.version_info.minor}"
        candidates: list[Path] = []
        try:
            user_sites = site.getusersitepackages()
            if isinstance(user_sites, str):
                candidates.append(Path(user_sites))
            else:
                candidates.extend(Path(path) for path in user_sites)
        except Exception:
            pass
        for home in (Path.home(), Path(__file__).resolve().parents[3]):
            candidates.append(
                home / "Library" / "Python" / version / "lib" / "python" / "site-packages"
            )
        for candidate in candidates:
            candidate_str = str(candidate)
            checked_paths.append(candidate_str)
            if candidate.exists() and candidate_str not in sys.path:
                sys.path.append(candidate_str)
        try:
            import ApplicationServices as AX
        except Exception as exc:
            raise RuntimeError(
                "Accessibility fallback unavailable "
                f"({type(exc).__name__}: {exc}; checked={checked_paths})"
            ) from exc

    pgrep = subprocess.run(
        ["pgrep", "-x", "Microsoft Outlook"],
        capture_output=True,
        text=True,
        timeout=5,
    )
    pids = [int(part) for part in pgrep.stdout.split() if part.strip().isdigit()]
    if not pids:
        raise RuntimeError("Outlook is not running")

    run_applescript(
        'tell application "Microsoft Outlook" to activate',
        timeout=5,
    )
    time.sleep(0.4)
    app = AX.AXUIElementCreateApplication(pids[0])
    windows = _ax_attr(AX, app, "AXWindows") or []
    if not windows:
        raise RuntimeError("No Outlook window available")
    window = list(windows)[0]
    window_title = _ax_attr(AX, window, "AXTitle") or ""
    tables = _ax_walk(AX, window, role="AXTable")
    for table in tables:
        rows = list(_ax_attr(AX, table, "AXRows") or _ax_attr(AX, table, "AXChildren") or [])
        if not rows:
            continue
        selected_rows = list(_ax_attr(AX, table, "AXSelectedRows") or [])
        message_like_rows = [
            row for row in rows if _looks_like_new_outlook_message_row(_ax_row_text(AX, row))
        ]
        if len(message_like_rows) < 2:
            continue
        selected_message_rows = [
            row
            for row in selected_rows
            if _looks_like_new_outlook_message_row(_ax_row_text(AX, row))
        ]
        if len(selected_message_rows) == 1:
            record = _parse_new_outlook_ax_row(
                _ax_row_text(AX, selected_message_rows[0]),
                window_title=window_title,
            )
            if record and _is_plausible_outlook_record(record):
                return record
        if (
            len(selected_message_rows) == len(message_like_rows)
            and len(message_like_rows) > 1
        ):
            record = _parse_new_outlook_ax_row(
                _ax_row_text(AX, message_like_rows[0]),
                window_title=window_title,
            )
            if record and _is_plausible_outlook_record(record):
                return record
        if len(selected_message_rows) > 1:
            raise RuntimeError(
                f"Select exactly one Outlook mail (Accessibility selected rows={len(selected_message_rows)})."
            )
    raise RuntimeError("No selected Outlook mail exposed through Accessibility")


def _legacy_outlook_fallback_enabled() -> bool:
    return os.environ.get("SAVE_MAIL_OUTLOOK_LEGACY_FALLBACK", "").strip().lower() in {
        "1",
        "true",
        "yes",
        "on",
    }


def _get_selected_mail_headers_legacy() -> list[dict[str, str]]:
    if not _outlook_ui_actions_enabled():
        raise RuntimeError("Legacy Outlook fallback requires SAVE_MAIL_OUTLOOK_UI_ACTIONS=1")

    script = r'''
tell application "Microsoft Outlook"
    activate
    set selectedMessages to {}
    try
        set selectedMessages to current messages
    end try
    try
        if selectedMessages is missing value then set selectedMessages to {}
    end try
    try
        if (count of selectedMessages) is 0 then set selectedMessages to selection
    end try
    try
        if (count of selectedMessages) is 0 then set selectedMessages to get selected objects
    end try

    set selectedMessageCount to 0
    set messageCandidates to {}
    try
        set selectedMessageCount to count of selectedMessages
        set messageCandidates to selectedMessages
    on error
        if selectedMessages is not missing value then
            set selectedMessageCount to 1
            set messageCandidates to {selectedMessages}
        else
            set selectedMessageCount to 0
            set messageCandidates to {}
        end if
    end try

    if selectedMessageCount is 0 then error "No selected Outlook mail."
    if selectedMessageCount > 1 then error "Select exactly one Outlook mail."

    set d to "|||"
    set recDelim to "|||NEXT|||"
    set allResults to {}

    repeat with msg in messageCandidates
        if msg is not missing value then
            set isMessage to false
            try
                if class of msg is message then set isMessage to true
            end try
            if isMessage then
                try
                    set msgSubject to ""
                    try
                        set msgSubject to subject of msg
                    end try

                    set senderDisplay to ""
                    set senderEmail to ""
                    try
                        set senderDisplay to sender of msg as string
                    end try
                    try
                        set senderEmail to email address of sender of msg as string
                    end try
                    if senderEmail is "" then
                        try
                            set senderEmail to senderDisplay
                        end try
                    end if

                    set msgDate to current date
                    try
                        set msgDate to time received of msg
                    on error
                        try
                            set msgDate to time sent of msg
                        end try
                    end try

                    set outlookId to ""
                    try
                        set outlookId to id of msg as string
                    end try

                    set acctName to "Microsoft Outlook"
                    set mboxName to "Outlook"
                    try
                        set mboxName to name of container of msg
                    end try
                    try
                        set acctName to name of account of container of msg
                    end try

                    set toList to {}
                    try
                        repeat with r in (to recipients of msg)
                            try
                                set end of toList to (email address of r as string)
                            on error
                                set end of toList to r as string
                            end try
                        end repeat
                    end try

                    set ccList to {}
                    try
                        repeat with r in (cc recipients of msg)
                            try
                                set end of ccList to (email address of r as string)
                            on error
                                set end of ccList to r as string
                            end try
                        end repeat
                    end try

                    set attCount to "0"
                    set attNames to {}
                    try
                        set attCount to count of attachments of msg
                        repeat with att in attachments of msg
                            try
                                set end of attNames to name of att
                            end try
                        end repeat
                    end try

                    set isFlagged to "false"
                    try
                        if flagged of msg then set isFlagged to "true"
                    end try

                    set msgHeaders to ""
                    try
                        set msgHeaders to headers of msg
                    end try
                    set hasCal to "false"
                    if msgHeaders contains "text/calendar" then set hasCal to "true"
                    if (attNames as string) contains ".ics" then set hasCal to "true"

                    set boxType to "selected"
                    if mboxName is not "" then
                        ignoring case
                            if mboxName contains "sent" or mboxName contains "verzonden" then
                                set boxType to "sent"
                            else if mboxName contains "deleted" or mboxName contains "trash" or mboxName contains "verwijder" or mboxName contains "prullenmand" then
                                set boxType to "deleted"
                            else if mboxName contains "archive" or mboxName contains "archiv" then
                                set boxType to "archive"
                            else if mboxName contains "inbox" or mboxName contains "postvak" then
                                set boxType to "inbox"
                            end if
                        end ignoring
                    end if

                    set end of allResults to (boxType & d & msgSubject & d & senderEmail & d & senderDisplay & d & ¬
                        (msgDate as string) & d & outlookId & d & acctName & d & mboxName & d & ¬
                        (toList as string) & d & (ccList as string) & d & ¬
                        attCount & d & (attNames as string) & d & isFlagged & d & hasCal & d & msgHeaders)
                end try
            end if
        end if
    end repeat

    if (count of allResults) is 0 then return ""
    set AppleScript's text item delimiters to recDelim
    return allResults as string
end tell
'''
    raw = run_applescript(script, timeout=60)
    mails: list[dict[str, str]] = []
    for record in raw.split("|||NEXT|||"):
        parsed = _parse_outlook_record(record.strip())
        if parsed and _is_plausible_outlook_record(parsed):
            if parsed.get("mailbox_type") == "selected":
                parsed["mailbox_type"] = classify_mailbox_type(
                    parsed.get("mailbox", ""), parsed.get("account", "")
                )
            mails.append(parsed)
    if not mails:
        raise RuntimeError("No selected Outlook mail metadata returned")
    return mails


def _get_selected_mail_body_legacy() -> str:
    if not _outlook_ui_actions_enabled():
        return ""

    script = r'''
tell application "Microsoft Outlook"
    activate
    set selectedMessages to {}
    try
        set selectedMessages to current messages
    end try
    try
        if selectedMessages is missing value then set selectedMessages to {}
    end try
    try
        if (count of selectedMessages) is 0 then set selectedMessages to selection
    end try
    try
        if (count of selectedMessages) is 0 then set selectedMessages to get selected objects
    end try

    set selectedMessageCount to 0
    set messageCandidates to {}
    try
        set selectedMessageCount to count of selectedMessages
        set messageCandidates to selectedMessages
    on error
        if selectedMessages is not missing value then
            set selectedMessageCount to 1
            set messageCandidates to {selectedMessages}
        else
            set selectedMessageCount to 0
            set messageCandidates to {}
        end if
    end try

    if selectedMessageCount is 0 then return ""
    if selectedMessageCount is not 1 then return ""
    set msg to item 1 of messageCandidates
    if msg is missing value then return ""
    set isMessage to false
    try
        if class of msg is message then set isMessage to true
    end try
    if not isMessage then return ""

    set rawContent to ""
    try
        set rawContent to content of msg
    end try
    if rawContent is "" then
        try
            set rawContent to plain text content of msg
        end try
    end if
    if (length of rawContent) > 50000 then
        return text 1 thru 50000 of rawContent
    end if
    return rawContent
end tell
'''
    try:
        return run_applescript(script, timeout=15)
    except Exception:
        return ""


def _load_new_outlook_selection() -> list[dict[str, str]]:
    global _LAST_NEW_OUTLOOK_RECORD
    try:
        record = _get_selected_mail_header_from_accessibility()
        _LAST_NEW_OUTLOOK_RECORD = record
        return [record]
    except RuntimeError as ax_exc:
        if _outlook_clipboard_fallback_enabled():
            try:
                record = _coerce_outlook_record_from_clipboard(
                    _copy_selected_outlook_message_via_clipboard()
                )
                if record and _is_plausible_outlook_record(record):
                    _LAST_NEW_OUTLOOK_RECORD = record
                    return [record]
            except RuntimeError as clipboard_exc:
                if _legacy_outlook_fallback_enabled():
                    return _get_selected_mail_headers_legacy()
                raise RuntimeError(
                    f"accessibility: {ax_exc}; clipboard: {clipboard_exc}"
                ) from clipboard_exc
        if _legacy_outlook_fallback_enabled():
            return _get_selected_mail_headers_legacy()
        raise RuntimeError(f"accessibility: {ax_exc}") from ax_exc


def _parse_outlook_record(record: str) -> dict[str, str] | None:
    parts = record.split("|||", 14)
    if len(parts) < 14:
        return None
    headers = _parse_headers(parts[14] if len(parts) > 14 else "")

    mailbox_type = parts[0] or classify_mailbox_type(parts[7], parts[6])
    subject = parts[1] or _pick_from_header(headers, ("subject",), "")
    sender_display = parts[3] or parts[2] or _extract_display_from_header(
        _pick_from_header(headers, ("from",), "")
    )
    sender_email = _extract_email(parts[2] or sender_display)
    if not sender_email:
        sender_email = _extract_email(_pick_from_header(headers, ("from",), ""))
    date_str = parts[4]
    all_headers = parts[14] if len(parts) > 14 else ""
    message_id = _message_id_from_parts(
        parts[5], all_headers, subject, sender_email, date_str
    )
    return {
        "subject": subject,
        "sender_email": sender_email,
        "sender_display": sender_display,
        "date_str": date_str,
        "message_id": message_id,
        "account": parts[6] or "Microsoft Outlook",
        "mailbox": parts[7] or "Inbox",
        "to": (parts[8].rstrip(", ") or _clean_address_list(_pick_from_header(headers, ("to",), ""))),
        "cc": (parts[9].rstrip(", ") or _clean_address_list(_pick_from_header(headers, ("cc",), ""))),
        "att_count": parts[10].strip() or "0",
        "att_names": parts[11].strip(),
        "is_flagged": _parse_bool(parts[12]),
        "has_calendar": _parse_bool(parts[13]),
        "all_headers": all_headers,
        "mailbox_type": mailbox_type,
        "body": "",
        "mail_client": "outlook",
    }


def _is_plausible_outlook_record(record: dict[str, str]) -> bool:
    subject = (record.get("subject") or "").strip()
    sender = (record.get("sender_email") or "").strip()
    date_str = (record.get("date_str") or "").strip()
    message_id = (record.get("message_id") or "").strip()
    if not any((subject, sender, date_str, message_id)):
        return False
    if subject == "" and sender == "" and message_id.startswith("outlook-"):
        return False
    return True


def get_selected_mail_headers() -> list[dict[str, str]]:
    """Fetch metadata for selected Outlook messages only."""
    if _is_new_outlook_enabled():
        return _load_new_outlook_selection()

    if _legacy_outlook_fallback_enabled():
        return _get_selected_mail_headers_legacy()

    script = r'''
tell application "Microsoft Outlook"
    activate
    set selectedMessages to {}
    try
        set selectedMessages to current messages
    end try
    try
        if selectedMessages is missing value then set selectedMessages to {}
    end try
    try
        if (count of selectedMessages) is 0 then set selectedMessages to selection
    end try
    try
        if (count of selectedMessages) is 0 then set selectedMessages to get selected objects
    end try

    set selectedMessageCount to 0
    set messageCandidates to {}
    try
        set selectedMessageCount to count of selectedMessages
        set messageCandidates to selectedMessages
    on error
        if selectedMessages is not missing value then
            set selectedMessageCount to 1
            set messageCandidates to {selectedMessages}
        else
            set selectedMessageCount to 0
            set messageCandidates to {}
        end if
    end try

    if selectedMessageCount is 0 then error "No selected Outlook mail."
    if selectedMessageCount > 1 then error "Select exactly one Outlook mail."

    set d to "|||"
    set recDelim to "|||NEXT|||"
    set allResults to {}

    repeat with msg in messageCandidates
        if msg is not missing value then
            set isMessage to false
            try
                if class of msg is message then set isMessage to true
            end try
            if isMessage then
                try
                    set msgSubject to ""
                    try
                        set msgSubject to subject of msg
                    end try

                    set senderDisplay to ""
                    set senderEmail to ""
                    try
                        set senderDisplay to sender of msg as string
                    end try
                    try
                        set senderEmail to email address of sender of msg as string
                    end try
                    if senderEmail is "" then
                        try
                            set senderEmail to senderDisplay
                        end try
                    end if

                    set msgDate to current date
                    try
                        set msgDate to time received of msg
                    on error
                        try
                            set msgDate to time sent of msg
                        end try
                    end try

                    set outlookId to ""
                    try
                        set outlookId to id of msg as string
                    end try

                    set acctName to "Microsoft Outlook"
                    set mboxName to "Outlook"
                    try
                        set mboxName to name of container of msg
                    end try
                    try
                        set acctName to name of account of container of msg
                    end try

                    set toList to {}
                    try
                        repeat with r in (to recipients of msg)
                            try
                                set end of toList to (email address of r as string)
                            on error
                                set end of toList to r as string
                            end try
                        end repeat
                    end try

                    set ccList to {}
                    try
                        repeat with r in (cc recipients of msg)
                            try
                                set end of ccList to (email address of r as string)
                            on error
                                set end of ccList to r as string
                            end try
                        end repeat
                    end try

                    set attCount to "0"
                    set attNames to {}
                    try
                        set attCount to count of attachments of msg
                        repeat with att in attachments of msg
                            try
                                set end of attNames to name of att
                            end try
                        end repeat
                    end try

                    set isFlagged to "false"
                    try
                        if flagged of msg then set isFlagged to "true"
                    end try

                    set msgHeaders to ""
                    try
                        set msgHeaders to headers of msg
                    end try
                    set hasCal to "false"
                    if msgHeaders contains "text/calendar" then set hasCal to "true"
                    if (attNames as string) contains ".ics" then set hasCal to "true"

                    set boxType to "selected"
                    if mboxName is not "" then
                        ignoring case
                            if mboxName contains "sent" or mboxName contains "verzonden" then
                                set boxType to "sent"
                            else if mboxName contains "deleted" or mboxName contains "trash" or mboxName contains "verwijder" or mboxName contains "prullenmand" then
                                set boxType to "deleted"
                            else if mboxName contains "archive" or mboxName contains "archiv" then
                                set boxType to "archive"
                            else if mboxName contains "inbox" or mboxName contains "postvak" then
                                set boxType to "inbox"
                            end if
                        end ignoring
                    end if

                    set end of allResults to (boxType & d & msgSubject & d & senderEmail & d & senderDisplay & d & ¬
                        (msgDate as string) & d & outlookId & d & acctName & d & mboxName & d & ¬
                        (toList as string) & d & (ccList as string) & d & ¬
                        attCount & d & (attNames as string) & d & isFlagged & d & hasCal & d & msgHeaders)
                end try
            end if
        end if
    end repeat

    if (count of allResults) is 0 then return ""
    set AppleScript's text item delimiters to recDelim
    return allResults as string
end tell
'''
    raw = run_applescript(script, timeout=60)
    mails: list[dict[str, str]] = []
    for record in raw.split("|||NEXT|||"):
        parsed = _parse_outlook_record(record.strip())
        if parsed and _is_plausible_outlook_record(parsed):
            if parsed.get("mailbox_type") == "selected":
                parsed["mailbox_type"] = classify_mailbox_type(
                    parsed.get("mailbox", ""), parsed.get("account", "")
                )
            mails.append(parsed)
    if not mails:
        raise RuntimeError("No selected Outlook mail metadata returned")
    return mails


def get_selected_mail_header() -> dict[str, str]:
    """Fetch metadata for the first selected Outlook message."""
    return get_selected_mail_headers()[0]


def fetch_mail_body(message_id: str, account: str = "", mailbox: str = "") -> str:
    """Fetch body for selected Outlook mail.

    Outlook fallback ids can be synthetic, so selected-mail lookup is more
    reliable than searching folders by id for the phase-1 selected-mail scope.
    """
    if _is_new_outlook_enabled():
        if _LAST_NEW_OUTLOOK_RECORD and (
            _clean_message_id(_LAST_NEW_OUTLOOK_RECORD.get("message_id", ""))
            == _clean_message_id(message_id)
        ):
            return _LAST_NEW_OUTLOOK_RECORD.get("body", "")
        try:
            ax_record = _get_selected_mail_header_from_accessibility()
            if (
                _clean_message_id(ax_record.get("message_id", ""))
                == _clean_message_id(message_id)
            ):
                return ax_record.get("body", "")
        except RuntimeError:
            pass
        if _outlook_clipboard_fallback_enabled():
            try:
                record = _coerce_outlook_record_from_clipboard(
                    _copy_selected_outlook_message_via_clipboard()
                )
                if record:
                    return record.get("body", "")
            except RuntimeError:
                if _LAST_NEW_OUTLOOK_RECORD and (
                    _clean_message_id(_LAST_NEW_OUTLOOK_RECORD.get("message_id", ""))
                    == _clean_message_id(message_id)
                ):
                    return _LAST_NEW_OUTLOOK_RECORD.get("body", "")
            if _legacy_outlook_fallback_enabled():
                return _get_selected_mail_body_legacy()
            return ""
        if _legacy_outlook_fallback_enabled():
            return _get_selected_mail_body_legacy()
        return ""

    script = r'''
tell application "Microsoft Outlook"
    activate
    set selectedMessages to {}
    try
        set selectedMessages to current messages
    end try
    try
        if selectedMessages is missing value then set selectedMessages to {}
    end try
    try
        if (count of selectedMessages) is 0 then set selectedMessages to selection
    end try
    try
        if (count of selectedMessages) is 0 then set selectedMessages to get selected objects
    end try

    set selectedMessageCount to 0
    set messageCandidates to {}
    try
        set selectedMessageCount to count of selectedMessages
        set messageCandidates to selectedMessages
    on error
        if selectedMessages is not missing value then
            set selectedMessageCount to 1
            set messageCandidates to {selectedMessages}
        else
            set selectedMessageCount to 0
            set messageCandidates to {}
        end if
    end try

    if selectedMessageCount is 0 then return ""
    if selectedMessageCount is not 1 then return ""
    set msg to item 1 of messageCandidates
    if msg is missing value then return ""
    set isMessage to false
    try
        if class of msg is message then set isMessage to true
    end try
    if not isMessage then return ""

    set rawContent to ""
    try
        set rawContent to content of msg
    end try
    if rawContent is "" then
        try
            set rawContent to plain text content of msg
        end try
    end if
    if (length of rawContent) > 50000 then
        return text 1 thru 50000 of rawContent
    end if
    return rawContent
end tell
'''
    try:
        return run_applescript(script, timeout=15)
    except Exception:
        return ""


def save_attachments(
    message_id: str,
    timestamp: str,
    account: str = "",
    mailbox: str = "",
    filename_prefix: str = "",
) -> list[str]:
    """Outlook attachment export is out of scope for the selected-mail MVP."""
    return []


def _archive_selected_outlook_message_via_menu() -> str:
    """Archive the currently selected Outlook message through the UI menu."""
    script = r'''
tell application "Microsoft Outlook"
    activate
end tell

delay 0.15

tell application "System Events"
    set outlookProcesses to (processes whose name is "Microsoft Outlook")
    if (count of outlookProcesses) is 0 then error "Outlook is not running"

    tell item 1 of outlookProcesses
        set frontmost to true
        delay 0.1

        set menuNames to {"Message", "Bericht"}
        set archiveNames to {"Archive", "Archiveren", "Archief", "Move to Archive", "Verplaatsen naar archief", "Verplaats naar archief"}

        repeat with menuName in menuNames
            try
                set messageMenu to menu bar item (menuName as text) of menu bar 1
                click messageMenu
                delay 0.1

                repeat with archiveName in archiveNames
                    try
                        click menu item (archiveName as text) of menu 1 of messageMenu
                        return "OK"
                    end try
                end repeat
            end try
        end repeat

        try
            click (first button of window 1 whose name contains "Archive")
            return "OK"
        end try
        try
            click (first button of window 1 whose description contains "Archive")
            return "OK"
        end try
        try
            click (first button of window 1 whose help contains "Archive")
            return "OK"
        end try
        try
            click (first button of window 1 whose name contains "Arch")
            return "OK"
        end try
        try
            click (first button of window 1 whose description contains "Arch")
            return "OK"
        end try

        error "Outlook Archive menu item not found"
    end tell
end tell
'''
    result = run_applescript(script, timeout=10)
    return result or "OK"


def _selected_outlook_message_matches(message_id: str) -> tuple[bool, str]:
    cleaned_expected = _clean_message_id(message_id)
    if not cleaned_expected:
        return False, "SKIPPED:no message id"
    try:
        selected = _get_selected_mail_header_from_accessibility()
    except RuntimeError as exc:
        return False, f"SKIPPED:outlook selection unavailable ({exc})"
    cleaned_selected = _clean_message_id(str(selected.get("message_id", "")))
    if cleaned_selected != cleaned_expected:
        return False, "SKIPPED:outlook selection changed"
    return True, ""


def archive_mail(message_id: str, account: str, mailbox: str = "") -> str:
    """Archive the currently selected Outlook mail after a successful save.

    New Outlook does not expose a reliable move/archive AppleScript API. For the
    selected-mail flow we verify that the current UI selection still matches the
    saved record, then trigger Outlook's Archive UI command.
    """
    cleaned_message_id = _clean_message_id(message_id)
    if _is_new_outlook_enabled() or cleaned_message_id.startswith("outlook-"):
        matches, reason = _selected_outlook_message_matches(message_id)
        if not matches:
            return reason
    try:
        return _archive_selected_outlook_message_via_menu()
    except Exception as exc:
        return f"ERROR:{exc}"


def fetch_all_mailboxes(since_days: int = 7) -> dict[str, list[dict[str, str]]]:
    """Batch Outlook sweep is out of scope for the selected-mail MVP."""
    return {"inbox": [], "sent": [], "deleted": []}
