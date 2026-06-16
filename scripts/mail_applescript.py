#!/usr/bin/env python3
"""Apple Mail AppleScript wrappers: fetch headers, body, attachments, archive.

Extracted from save_mail.py for reuse and maintainability.
"""

from __future__ import annotations

import re
import shutil
import subprocess
import sys
from pathlib import Path

import brain_lib

ATTACHMENTS_DIR = brain_lib.cfg.vault_files

# Archive folder per account
ARCHIVE_FOLDERS: dict[str, str] = {
    "iCloud": "Archive",
    "McCoy": "Archive",
    "Proton": "Archive",
    "Gmail": "Archive",
    "viggomulders": "Archiveren",
    "ricardo001": "Archiveren",
    "1990_kees": "Archiveren",
}


def is_mail_running() -> bool:
    """Check if Mail.app is running."""
    try:
        result = subprocess.run(["pgrep", "-x", "Mail"], capture_output=True, timeout=5)
        return result.returncode == 0
    except Exception:
        return False


# Mailbox name mapping per type — varies per account/provider/locale
INBOX_MAILBOX_NAMES = ("INBOX", "Inbox", "Postvak IN")
DELETED_MAILBOX_NAMES = (
    "Deleted Messages",
    "Deleted Items",
    "Trash",
    "Bin",
    "Verwijderde items",
    "Prullenmand",
)
SENT_MAILBOX_NAMES = (
    "Sent Messages",
    "Sent Mail",
    "Sent",
    "Sent Items",
    "Verzonden items",
)
ARCHIVE_MAILBOX_NAMES = tuple(
    dict.fromkeys(("Archive", "Archiveren", *ARCHIVE_FOLDERS.values()))
)
CONVERSATION_EXACT_TIMEOUT_SECONDS = 60
CONVERSATION_SUBJECT_TIMEOUT_SECONDS = 12


def classify_mailbox_type(mailbox_name: str, account: str = "") -> str:
    """Classify a Mail.app mailbox name into the small set save_mail cares about."""
    normalized = mailbox_name.strip().lower()
    archive_name = ARCHIVE_FOLDERS.get(account, "").strip().lower()
    if normalized in {n.lower() for n in INBOX_MAILBOX_NAMES}:
        return "inbox"
    if normalized in {n.lower() for n in SENT_MAILBOX_NAMES}:
        return "sent"
    if normalized in {n.lower() for n in DELETED_MAILBOX_NAMES}:
        return "deleted"
    if archive_name and normalized == archive_name:
        return "archive"
    if normalized in {n.lower() for n in ARCHIVE_MAILBOX_NAMES}:
        return "archive"
    return "other"


def _applescript_string(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def _applescript_list(values: list[str]) -> str:
    return "{" + ", ".join(f'"{_applescript_string(v)}"' for v in values) + "}"


def _parse_headers(raw_headers: str) -> dict[str, str]:
    headers: dict[str, str] = {}
    current_lines: list[str] = []
    for raw_line in (raw_headers or "").replace("\r\n", "\n").split("\n"):
        if not raw_line:
            continue
        if raw_line[:1] in (" ", "\t") and current_lines:
            current_lines[-1] = current_lines[-1] + " " + raw_line.strip()
        else:
            current_lines.append(raw_line.strip())

    for line in current_lines:
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        headers[key.strip().lower()] = value.strip()
    return headers


def _clean_message_id(value: str) -> str:
    return (value or "").strip().strip("<>").strip()


def _message_id_from_headers(raw_headers: str) -> str:
    return _clean_message_id(_parse_headers(raw_headers).get("message-id", ""))


def _parse_mail_record(record: str) -> dict[str, str] | None:
    parts = record.split("|||", 14)
    if len(parts) < 13:
        return None
    mailbox_type = parts[0]
    all_headers = parts[14] if len(parts) > 14 else ""
    message_id = _clean_message_id(parts[5]) or _message_id_from_headers(all_headers)
    return {
        "subject": parts[1],
        "sender_email": parts[2].strip().lower(),
        "sender_display": parts[3],
        "date_str": parts[4],
        "message_id": message_id,
        "account": parts[6],
        "mailbox": parts[7],
        "to": parts[8].rstrip(", "),
        "cc": parts[9].rstrip(", "),
        "att_count": parts[10].strip(),
        "att_names": parts[11].strip(),
        "is_flagged": parts[12].strip() == "true",
        "has_calendar": parts[13].strip() == "true" if len(parts) > 13 else False,
        "all_headers": all_headers,
        "mailbox_type": mailbox_type,
        "body": "",
    }


def fetch_all_mailboxes(since_days: int = 7) -> dict[str, list[dict[str, str]]]:
    """Fetch inbox+sent+deleted from all accounts in one pass (7 calls instead of 21).

    Returns dict with keys "inbox", "sent", "deleted", each a list of mail dicts.
    """
    if not is_mail_running():
        return {"inbox": [], "sent": [], "deleted": []}

    result: dict[str, list[dict[str, str]]] = {"inbox": [], "sent": [], "deleted": []}
    accounts_list = list(ARCHIVE_FOLDERS.keys())

    for acct_name in accounts_list:
        try:
            mails = _fetch_all_from_account(acct_name, since_days)
            for m in mails:
                mtype = m.get("mailbox_type", "inbox")
                if mtype in result:
                    result[mtype].append(m)
        except Exception as e:
            print(f"⚠️  {acct_name}: {e}", file=sys.stderr, flush=True)
            continue

    return result


def _fetch_all_from_account(
    acct_name: str, since_days: int = 7
) -> list[dict[str, str]]:
    """Fetch inbox+sent+deleted from one account in a single AppleScript call."""
    inbox_names = '", "'.join(INBOX_MAILBOX_NAMES)
    sent_names = '", "'.join(SENT_MAILBOX_NAMES)
    deleted_names = '", "'.join(DELETED_MAILBOX_NAMES)

    date_line = ""
    inbox_filter = "whose read status is true"
    other_filter = ""
    if since_days > 0:
        date_line = f"set cutoffDate to (current date) - ({since_days} * days)"
        inbox_filter = "whose read status is true and date received > cutoffDate"
        other_filter = "whose date received > cutoffDate"

    # Generate mailbox extraction blocks (DRY: one template, 3 instances)
    mailbox_specs = [
        ("inbox", inbox_names, inbox_filter),
        ("sent", sent_names, other_filter),
        ("deleted", deleted_names, other_filter),
    ]
    mailbox_blocks = []
    for type_tag, names, filt in mailbox_specs:
        filt_clause = f" {filt}" if filt else ""
        mailbox_blocks.append(f'''
    -- {type_tag.upper()}
    set targetBox to missing value
    repeat with boxName in {{"{names}"}}
        try
            set targetBox to mailbox boxName of targetAcct
            exit repeat
        end try
    end repeat
    if targetBox is not missing value then
        try
            set msgs to (messages of targetBox{filt_clause})
            repeat with msg in msgs
                try
                    set msgSubject to subject of msg
                    set msgSender to sender of msg
                    set msgDate to date received of msg
                    set msgId to message id of msg
                    set mboxName to name of targetBox
                    set toList to {{}}
                    repeat with r in (to recipients of msg)
                        set end of toList to address of r
                    end repeat
                    set ccList to {{}}
                    repeat with r in (cc recipients of msg)
                        set end of ccList to address of r
                    end repeat
                    set senderEmail to msgSender
                    if msgSender contains "<" then
                        set senderEmail to text ((offset of "<" in msgSender) + 1) thru ((offset of ">" in msgSender) - 1) of msgSender
                    end if
                    set attCount to count of mail attachments of msg
                    set attNames to {{}}
                    repeat with att in mail attachments of msg
                        set end of attNames to name of att
                    end repeat
                    set isFlagged to "false"
                    if flagged status of msg then set isFlagged to "true"
                    set hasCal to "false"
                    if "{type_tag}" is "deleted" then
                        try
                            if source of msg contains "text/calendar" then set hasCal to "true"
                        end try
                    end if
                    set end of allResults to ("{type_tag}" & d & msgSubject & d & senderEmail & d & msgSender & d & (msgDate as string) & d & msgId & d & "{acct_name}" & d & mboxName & d & (toList as string) & d & (ccList as string) & d & attCount & d & (attNames as string) & d & isFlagged & d & hasCal)
                end try
            end repeat
        end try
    end if''')

    all_blocks = "\n".join(mailbox_blocks)

    script = f'''
tell application "Mail"
    set targetAcct to missing value
    repeat with acct in accounts
        if name of acct is "{acct_name}" then
            set targetAcct to acct
            exit repeat
        end if
    end repeat
    if targetAcct is missing value then return ""

    {date_line}

    set d to "|||"
    set recDelim to "|||NEXT|||"
    set allResults to {{}}
{all_blocks}

    if (count of allResults) is 0 then return ""
    set AppleScript's text item delimiters to recDelim
    return allResults as string
end tell
'''

    try:
        raw = run_applescript(script, timeout=60)
    except Exception as e:
        print(f"⚠️  {acct_name}/combined: {e}", file=sys.stderr, flush=True)
        return []

    if not raw or raw.strip() == "":
        return []

    results: list[dict[str, str]] = []
    for record in raw.split("|||NEXT|||"):
        record = record.strip()
        if not record:
            continue
        parsed = _parse_mail_record(record)
        if parsed:
            results.append(parsed)

    return results


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
        raise TimeoutError(f"AppleScript timed out after {timeout}s") from e
    if result.returncode != 0:
        raise RuntimeError(f"AppleScript error: {result.stderr.strip()}")
    return result.stdout.strip()


def get_selected_mail_headers() -> list[dict[str, str]]:
    """Fetch metadata for selected messages only (no body)."""
    script = """
tell application "Mail"
    set selMsgs to selection
    if (count of selMsgs) is 0 then
        error "No mail selected"
    end if
    set d to "|||"
    set recDelim to "|||NEXT|||"
    set allResults to {}

    repeat with msg in selMsgs
        try
            set msgSubject to subject of msg
            set msgSender to sender of msg
            set msgDate to date received of msg
            set msgId to message id of msg
            set acctName to name of account of mailbox of msg
            set mboxName to name of mailbox of msg

            set toList to {}
            repeat with r in (to recipients of msg)
                set end of toList to address of r
            end repeat
            set ccList to {}
            repeat with r in (cc recipients of msg)
                set end of ccList to address of r
            end repeat

            set senderEmail to msgSender
            if msgSender contains "<" then
                set senderEmail to text ((offset of "<" in msgSender) + 1) thru ((offset of ">" in msgSender) - 1) of msgSender
            end if

            set attCount to count of mail attachments of msg
            set attNames to {}
            repeat with att in mail attachments of msg
                set end of attNames to name of att
            end repeat
            set isFlagged to "false"
            if flagged status of msg then set isFlagged to "true"
            set msgHeaders to ""
            try
                set msgHeaders to all headers of msg
            end try
            set hasCal to "false"
            if msgHeaders contains "text/calendar" then set hasCal to "true"

            set end of allResults to ("selected" & d & msgSubject & d & senderEmail & d & msgSender & d & ¬
                (msgDate as string) & d & msgId & d & acctName & d & mboxName & d & ¬
                (toList as string) & d & (ccList as string) & d & ¬
                attCount & d & (attNames as string) & d & isFlagged & d & hasCal & d & msgHeaders)
        end try
    end repeat

    if (count of allResults) is 0 then return ""
    set AppleScript's text item delimiters to recDelim
    return allResults as string
end tell
"""
    raw = run_applescript(script, timeout=60)
    mails: list[dict[str, str]] = []
    for record in raw.split("|||NEXT|||"):
        record = record.strip()
        if not record:
            continue
        parsed = _parse_mail_record(record)
        if parsed:
            parsed["mailbox_type"] = classify_mailbox_type(
                parsed.get("mailbox", ""), parsed.get("account", "")
            )
            mails.append(parsed)
    if not mails:
        raise RuntimeError("No selected mail metadata returned")
    return mails


def get_selected_mail_header() -> dict[str, str]:
    """Fetch metadata for the first selected message."""
    return get_selected_mail_headers()[0]


def _conversation_mailbox_specs(account: str, selected_mailbox: str = "") -> list[tuple[str, str]]:
    specs: list[tuple[str, str]] = []
    if selected_mailbox:
        specs.append((selected_mailbox, classify_mailbox_type(selected_mailbox, account)))
    specs.extend((name, "inbox") for name in INBOX_MAILBOX_NAMES)
    specs.extend((name, "sent") for name in SENT_MAILBOX_NAMES)
    archive_name = ARCHIVE_FOLDERS.get(account, "Archive")
    specs.append((archive_name, "archive"))
    specs.extend((name, "archive") for name in ARCHIVE_MAILBOX_NAMES)

    seen: set[str] = set()
    unique_specs: list[tuple[str, str]] = []
    for name, mailbox_type in specs:
        key = name.strip().lower()
        if not key or key in seen:
            continue
        seen.add(key)
        unique_specs.append((name, mailbox_type))
    return unique_specs


def fetch_conversation_candidates(
    selected_mail: dict[str, str],
    message_ids: list[str],
    subject_hint: str = "",
    *,
    include_subject: bool = False,
    search_other_accounts: bool = True,
    since_days: int = 180,
    max_subject_matches: int = 40,
) -> list[dict[str, str]]:
    """Fetch possible messages from the selected mail's conversation.

    The caller decides which candidates are high-confidence; this function only
    gathers plausible records from Inbox, Sent, and Archive mailboxes.
    """
    if not is_mail_running():
        return []

    selected_account = selected_mail.get("account", "")
    accounts = []
    if selected_account:
        accounts.append(selected_account)
    accounts.extend(a for a in ARCHIVE_FOLDERS if a not in accounts)

    subject_hint = subject_hint.strip()
    use_subject = include_subject and len(subject_hint) >= 3
    message_ids = [mid for mid in message_ids if _clean_message_id(mid)]
    if not message_ids and not use_subject:
        return []
    by_message_id: dict[str, dict[str, str]] = {}

    def add_records(records: list[dict[str, str]]) -> None:
        for mail in records:
            key = mail.get("message_id", "").strip().lower()
            if key and key not in by_message_id:
                by_message_id[key] = mail

    if selected_account:
        add_records(
            _fetch_conversation_candidates_from_account(
                account=selected_account,
                selected_mailbox=selected_mail.get("mailbox", ""),
                message_ids=message_ids,
                subject_hint=subject_hint,
                include_subject=use_subject,
                since_days=since_days,
                max_subject_matches=max_subject_matches,
            )
        )
        if not use_subject:
            return list(by_message_id.values())
        if len(by_message_id) > 1:
            return list(by_message_id.values())
        if not search_other_accounts:
            return list(by_message_id.values())

    for account in accounts:
        if account == selected_account:
            continue
        raw = _fetch_conversation_candidates_from_account(
            account=account,
            selected_mailbox="",
            message_ids=message_ids,
            subject_hint=subject_hint,
            include_subject=use_subject,
            since_days=since_days,
            max_subject_matches=max_subject_matches,
        )
        add_records(raw)

    return list(by_message_id.values())


def _fetch_conversation_candidates_from_account(
    *,
    account: str,
    selected_mailbox: str,
    message_ids: list[str],
    subject_hint: str,
    include_subject: bool,
    since_days: int,
    max_subject_matches: int,
) -> list[dict[str, str]]:
    mailbox_specs = _conversation_mailbox_specs(account, selected_mailbox)
    mailbox_list = (
        "{"
        + ", ".join(
            f'{{"{_applescript_string(name)}", "{mailbox_type}"}}'
            for name, mailbox_type in mailbox_specs
        )
        + "}"
    )
    target_ids = _applescript_list(
        list(dict.fromkeys(_clean_message_id(mid) for mid in message_ids if _clean_message_id(mid)))
    )
    subject_literal = _applescript_string(subject_hint)
    subject_flag = "true" if include_subject and subject_hint else "false"

    script = f'''
on formatMailRecord(msg, acctName, boxType, boxName, d)
    tell application "Mail"
        set msgSubject to subject of msg
        set msgSender to sender of msg
        set msgDate to date received of msg
        set msgId to message id of msg
        set toList to {{}}
        repeat with r in (to recipients of msg)
            set end of toList to address of r
        end repeat
        set ccList to {{}}
        repeat with r in (cc recipients of msg)
            set end of ccList to address of r
        end repeat
        set senderEmail to msgSender
        if msgSender contains "<" then
            set senderEmail to text ((offset of "<" in msgSender) + 1) thru ((offset of ">" in msgSender) - 1) of msgSender
        end if
        set attCount to count of mail attachments of msg
        set attNames to {{}}
        repeat with att in mail attachments of msg
            set end of attNames to name of att
        end repeat
        set isFlagged to "false"
        if flagged status of msg then set isFlagged to "true"
        set hasCal to "false"
        try
            if source of msg contains "text/calendar" then set hasCal to "true"
        end try
        set msgHeaders to ""
        try
            set msgHeaders to all headers of msg
        end try
        return boxType & d & msgSubject & d & senderEmail & d & msgSender & d & (msgDate as string) & d & msgId & d & acctName & d & boxName & d & (toList as string) & d & (ccList as string) & d & attCount & d & (attNames as string) & d & isFlagged & d & hasCal & d & msgHeaders
    end tell
end formatMailRecord

tell application "Mail"
    set targetAcct to missing value
    repeat with acct in accounts
        if name of acct is "{_applescript_string(account)}" then
            set targetAcct to acct
            exit repeat
        end if
    end repeat
    if targetAcct is missing value then return ""

    set d to "|||"
    set recDelim to "|||NEXT|||"
    set allResults to {{}}
    set mailboxSpecs to {mailbox_list}
    set targetIds to {target_ids}
    set shouldSearchSubject to {subject_flag}
    set subjectHint to "{subject_literal}"
    set cutoffDate to (current date) - ({since_days} * days)
    set subjectLimit to {max_subject_matches}

    repeat with mailboxSpec in mailboxSpecs
        set boxName to item 1 of mailboxSpec
        set boxType to item 2 of mailboxSpec
        set targetBox to missing value
        try
            set targetBox to mailbox boxName of targetAcct
        end try
        if targetBox is not missing value then
            repeat with targetId in targetIds
                try
                    set msgs to (messages of targetBox whose message id is (targetId as string))
                    repeat with msg in msgs
                        set end of allResults to my formatMailRecord(msg, "{_applescript_string(account)}", boxType, boxName, d)
                    end repeat
                end try
            end repeat

            if shouldSearchSubject is true then
                try
                    set subjectMatches to (messages of targetBox whose subject contains subjectHint and date received > cutoffDate)
                    set subjectCount to 0
                    repeat with msg in subjectMatches
                        if subjectCount >= subjectLimit then exit repeat
                        set end of allResults to my formatMailRecord(msg, "{_applescript_string(account)}", boxType, boxName, d)
                        set subjectCount to subjectCount + 1
                    end repeat
                end try
            end if
        end if
    end repeat

    if (count of allResults) is 0 then return ""
    set AppleScript's text item delimiters to recDelim
    return allResults as string
end tell
'''
    try:
        timeout = (
            CONVERSATION_SUBJECT_TIMEOUT_SECONDS
            if include_subject
            else CONVERSATION_EXACT_TIMEOUT_SECONDS
        )
        raw = run_applescript(script, timeout=timeout)
    except Exception as e:
        print(f"⚠️  {account}/conversation: {e}", file=sys.stderr, flush=True)
        return []
    if not raw:
        return []

    results: list[dict[str, str]] = []
    for record in raw.split("|||NEXT|||"):
        parsed = _parse_mail_record(record.strip())
        if parsed:
            results.append(parsed)
    return results


def fetch_mail_body(message_id: str, account: str = "", mailbox: str = "") -> str:
    """Fetch mail body by message-id. Uses known account/mailbox for fast path."""
    message_id = _clean_message_id(message_id)
    if not message_id:
        return ""

    if account and mailbox:
        script = f'''
tell application "Mail"
    try
        set targetAcct to first account whose name is "{account}"
        set targetBox to mailbox "{mailbox}" of targetAcct
        set msgs to (messages of targetBox whose message id is "{message_id}")
        if (count of msgs) > 0 then
            set rawContent to content of item 1 of msgs
            if (length of rawContent) > 50000 then
                return text 1 thru 50000 of rawContent
            end if
            return rawContent
        end if
    end try
    return ""
end tell
'''
    else:
        script = f'''
tell application "Mail"
    repeat with acct in accounts
        repeat with mbox in mailboxes of acct
            try
                set msgs to (messages of mbox whose message id is "{message_id}")
                if (count of msgs) > 0 then
                    set rawContent to content of item 1 of msgs
                    if (length of rawContent) > 50000 then
                        return text 1 thru 50000 of rawContent
                    end if
                    return rawContent
                end if
            end try
        end repeat
    end repeat
    return ""
end tell
'''
    try:
        return run_applescript(script, timeout=15)
    except Exception:
        return ""


def _attachment_original_name(saved_name: str, filename_prefix: str = "") -> str:
    if filename_prefix and saved_name.startswith(f"{filename_prefix}-"):
        return saved_name[len(filename_prefix) + 1 :]
    return saved_name.split("-", 2)[-1] if saved_name.count("-") >= 2 else saved_name


def _attachment_wikilink(saved_name: str, filename_prefix: str = "") -> str:
    display = _attachment_original_name(saved_name, filename_prefix)
    return f"[[{saved_name}|{display}]]"


def _unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    counter = 2
    while True:
        candidate = path.with_name(f"{path.stem}-{counter}{path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def save_attachments(
    message_id: str,
    timestamp: str,
    account: str = "",
    mailbox: str = "",
    filename_prefix: str = "",
) -> list[str]:
    """Save all attachments of a mail to vault attachments dir. Returns wikilinks."""
    message_id = _clean_message_id(message_id)
    if not message_id:
        return []

    staging_dir = brain_lib.artifact_dir("mail-attachment", timestamp, create=True)
    attachment_prefix = filename_prefix.strip() or timestamp

    if account and mailbox:
        find_block = f'''
    set targetAcct to missing value
    repeat with acct in accounts
        if name of acct is "{account}" then
            set targetAcct to acct
            exit repeat
        end if
    end repeat
    if targetAcct is missing value then return "ERROR:account not found"
    set srcBox to mailbox "{mailbox}" of targetAcct
    set msgs to (messages of srcBox whose message id is "{message_id}")
    if (count of msgs) is 0 then return "ERROR:Mail not found"
    set targetMsg to item 1 of msgs'''
    else:
        find_block = f'''
    set targetMsg to missing value
    repeat with acct in accounts
        repeat with mbox in mailboxes of acct
            try
                set msgs to (messages of mbox whose message id is "{message_id}")
                if (count of msgs) > 0 then
                    set targetMsg to item 1 of msgs
                    exit repeat
                end if
            end try
        end repeat
        if targetMsg is not missing value then exit repeat
    end repeat
    if targetMsg is missing value then return "ERROR:Mail not found"'''

    script = f'''
tell application "Mail"
{find_block}

    set attCount to count of mail attachments of targetMsg
    if attCount is 0 then return "NONE"

    set savedFiles to {{}}
    set attDir to POSIX path of "{staging_dir}"
    repeat with att in mail attachments of targetMsg
        set attName to name of att
        set saveName to "{_applescript_string(attachment_prefix)}-" & attName
        set savePath to attDir & "/" & saveName
        try
            save att in POSIX file savePath
            set end of savedFiles to saveName
        end try
    end repeat

    set AppleScript's text item delimiters to "|||"
    return savedFiles as string
end tell
'''
    try:
        result = run_applescript(script, timeout=30)
        if result == "NONE" or result.startswith("ERROR:"):
            return []
        names = [n.strip() for n in result.split("|||") if n.strip()]
        # Filter inline images (email signatures, logos, tracking pixels)
        INLINE_PATTERNS = [
            r"image00\d\.(png|jpg|jpeg|gif)",
            r"logo[s_-].*\.(png|jpg|jpeg|gif|svg)",
            r"footer[s_-]?.*\.(png|jpg|jpeg|gif|svg)",
            r"banner[s_-]?.*\.(png|jpg|jpeg|gif)",
            r"spacer.*\.(png|gif)",
            r"pixel\.(png|gif)",
            r"icon[s_-]?.*\.(png|jpg|jpeg|gif)",
            r"signature.*\.(png|jpg|jpeg|gif)",
            r"header[s_-]?.*\.(png|jpg|jpeg|gif)",
        ]
        filtered = []
        for name in names:
            orig_name = _attachment_original_name(name, attachment_prefix)
            if any(re.match(p, orig_name, re.IGNORECASE) for p in INLINE_PATTERNS):
                (staging_dir / name).unlink(missing_ok=True)
                continue
            att_path = staging_dir / name
            if att_path.exists() and att_path.stat().st_size < 10240:
                att_ext = att_path.suffix.lower()
                if att_ext in (".png", ".jpg", ".jpeg", ".gif", ".svg"):
                    att_path.unlink()
                    continue
            dest_path = brain_lib.canonical_artifact_path(
                name, timestamp, create_parent=True
            )
            if att_path.exists() and att_path.resolve() != dest_path.resolve():
                dest_path = _unique_path(dest_path)
                shutil.move(str(att_path), dest_path)
                filtered.append(dest_path.name)
            else:
                filtered.append(name)
        return [_attachment_wikilink(name, attachment_prefix) for name in filtered]
    except Exception:
        return []


def archive_mail(message_id: str, account: str, mailbox: str = "") -> str:
    """Archive the mail in its account's archive folder."""
    message_id = _clean_message_id(message_id)
    if not message_id:
        return "SKIPPED:no message id"

    archive_folder = ARCHIVE_FOLDERS.get(account, "Archive")
    if mailbox:
        script = f'''
tell application "Mail"
    set targetAcct to missing value
    repeat with acct in accounts
        if name of acct is "{account}" then
            set targetAcct to acct
            exit repeat
        end if
    end repeat
    if targetAcct is missing value then return "ERROR:account not found"

    set srcBox to mailbox "{mailbox}" of targetAcct
    set msgs to (messages of srcBox whose message id is "{message_id}")
    if (count of msgs) is 0 then return "ERROR:not found"
    set targetMsg to item 1 of msgs

    set archiveBox to mailbox "{archive_folder}" of targetAcct
    set mailbox of targetMsg to archiveBox
    return "OK"
end tell
'''
    else:
        script = f'''
tell application "Mail"
    set targetMsg to missing value
    set targetAcct to missing value
    repeat with acct in accounts
        if name of acct is "{account}" then
            set targetAcct to acct
            repeat with mbox in mailboxes of acct
                try
                    set msgs to (messages of mbox whose message id is "{message_id}")
                    if (count of msgs) > 0 then
                        set targetMsg to item 1 of msgs
                        exit repeat
                    end if
                end try
            end repeat
        end if
        if targetMsg is not missing value then exit repeat
    end repeat

    if targetMsg is missing value then return "ERROR:not found"
    if targetAcct is missing value then return "ERROR:account not found"

    set archiveBox to mailbox "{archive_folder}" of targetAcct
    set mailbox of targetMsg to archiveBox
    return "OK"
end tell
'''
    try:
        result = run_applescript(script, timeout=15)
        return result
    except Exception as e:
        return f"ERROR:{e}"
