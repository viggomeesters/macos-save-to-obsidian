#!/usr/bin/env python3
"""Microsoft Outlook AppleScript wrappers for selected-mail save support."""

from __future__ import annotations

import hashlib
import re
import subprocess

from mail_applescript import _parse_headers, classify_mailbox_type


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
            raise RuntimeError("No selected Outlook mail.")
        if "Select exactly one Outlook mail" in stderr:
            raise RuntimeError("Select exactly one Outlook mail.")
        if "-1743" in stderr or "not authorized" in stderr.lower():
            raise RuntimeError(
                "Microsoft Outlook Automation permission is missing for osascript"
            )
        raise RuntimeError(f"Outlook AppleScript error: {stderr}")
    return result.stdout.strip()


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


def _parse_outlook_record(record: str) -> dict[str, str] | None:
    parts = record.split("|||", 14)
    if len(parts) < 14:
        return None
    mailbox_type = parts[0] or classify_mailbox_type(parts[7], parts[6])
    subject = parts[1]
    sender_display = parts[3] or parts[2]
    sender_email = _extract_email(parts[2] or sender_display)
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
        "to": parts[8].rstrip(", "),
        "cc": parts[9].rstrip(", "),
        "att_count": parts[10].strip() or "0",
        "att_names": parts[11].strip(),
        "is_flagged": _parse_bool(parts[12]),
        "has_calendar": _parse_bool(parts[13]),
        "all_headers": all_headers,
        "mailbox_type": mailbox_type,
        "body": "",
        "mail_client": "outlook",
    }


def get_selected_mail_headers() -> list[dict[str, str]]:
    """Fetch metadata for selected Outlook messages only."""
    script = r'''
tell application "Microsoft Outlook"
    set selMsgs to selection
    if selMsgs is missing value then error "No selected Outlook mail."
    try
        set selCount to count of selMsgs
    on error
        set selMsgs to {selMsgs}
        set selCount to 1
    end try
    if selCount is 0 then error "No selected Outlook mail."
    if selCount > 1 then error "Select exactly one Outlook mail."

    set d to "|||"
    set recDelim to "|||NEXT|||"
    set allResults to {}

    repeat with msg in selMsgs
        try
            set msgSubject to ""
            try
                set msgSubject to subject of msg
            end try

            set senderDisplay to ""
            set senderEmail to ""
            try
                set senderDisplay to sender of msg as string
                set senderEmail to senderDisplay
            end try
            try
                set senderEmail to email address of sender of msg as string
            end try

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
                        set end of toList to address of r
                    on error
                        set end of toList to r as string
                    end try
                end repeat
            end try

            set ccList to {}
            try
                repeat with r in (cc recipients of msg)
                    try
                        set end of ccList to address of r
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
                set lowerBox to my lowercaseText(mboxName)
                if lowerBox contains "sent" or lowerBox contains "verzonden" then
                    set boxType to "sent"
                else if lowerBox contains "deleted" or lowerBox contains "trash" or lowerBox contains "verwijder" or lowerBox contains "prullenmand" then
                    set boxType to "deleted"
                else if lowerBox contains "archive" or lowerBox contains "archiv" then
                    set boxType to "archive"
                else if lowerBox contains "inbox" or lowerBox contains "postvak" then
                    set boxType to "inbox"
                end if
            end if

            set end of allResults to (boxType & d & msgSubject & d & senderEmail & d & senderDisplay & d & ¬
                (msgDate as string) & d & outlookId & d & acctName & d & mboxName & d & ¬
                (toList as string) & d & (ccList as string) & d & ¬
                attCount & d & (attNames as string) & d & isFlagged & d & hasCal & d & msgHeaders)
        end try
    end repeat

    if (count of allResults) is 0 then return ""
    set AppleScript's text item delimiters to recDelim
    return allResults as string
end tell

on lowercaseText(theText)
    set upperChars to "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    set lowerChars to "abcdefghijklmnopqrstuvwxyz"
    set outText to ""
    repeat with i from 1 to length of theText
        set c to character i of theText
        set pos to offset of c in upperChars
        if pos > 0 then
            set outText to outText & character pos of lowerChars
        else
            set outText to outText & c
        end if
    end repeat
    return outText
end lowercaseText
'''
    raw = run_applescript(script, timeout=60)
    mails: list[dict[str, str]] = []
    for record in raw.split("|||NEXT|||"):
        parsed = _parse_outlook_record(record.strip())
        if parsed:
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
    script = r'''
tell application "Microsoft Outlook"
    set selMsgs to selection
    if selMsgs is missing value then return ""
    try
        if (count of selMsgs) is 0 then return ""
        set msg to item 1 of selMsgs
    on error
        set msg to selMsgs
    end try

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


def archive_mail(message_id: str, account: str, mailbox: str = "") -> str:
    """Outlook archiving is out of scope for the selected-mail MVP."""
    return "SKIPPED:outlook archive not implemented"


def fetch_all_mailboxes(since_days: int = 7) -> dict[str, list[dict[str, str]]]:
    """Batch Outlook sweep is out of scope for the selected-mail MVP."""
    return {"inbox": [], "sent": [], "deleted": []}
