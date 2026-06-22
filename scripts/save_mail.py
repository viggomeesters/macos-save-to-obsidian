#!/usr/bin/env python3
"""Save the currently selected Apple Mail or Microsoft Outlook message to the vault.

Designed for speed: AppleScript fetch → SQLite dedup → entity resolve → vault write.
Called directly or via Raycast Script Command.

Usage:
    python3 save_mail.py              # Save selected mail; flagged mail creates a task
    python3 save_mail.py --client outlook
    python3 save_mail.py --task       # Force follow-up task for selected mail
    python3 save_mail.py --no-archive # Save without archiving
"""

from __future__ import annotations

import importlib
import hashlib
import json
import re
import sys
from datetime import datetime
from contextlib import contextmanager, nullcontext
from pathlib import Path
from types import ModuleType
from typing import Any, Iterator

import yaml

REPO_SHARED = Path(__file__).resolve().parents[1] / "shared"
if str(REPO_SHARED) not in sys.path:
    sys.path.insert(0, str(REPO_SHARED))
import brain_lib

from entity_resolver import (
    SELF_EMAILS,
    resolve_entity,
    resolve_entity_details,
    suggest_topics,
)
from mail_project_rules import MailProjectMatch, resolve_mail_project
import mail_applescript
from mail_applescript import (
    archive_mail,
    classify_mailbox_type,
    fetch_all_mailboxes,
    fetch_conversation_candidates,
    fetch_mail_body,
    get_selected_mail_header,
    get_selected_mail_headers,
    is_mail_running,
    save_attachments,
)

# Re-export for backward compatibility (save_message.py imports from here)
from entity_resolver import SELF_EMAILS as SELF_EMAILS  # noqa: F811
import vault_note_writer as vnw
from vault_note_writer import escape_title  # noqa: F401

VAULT_ROOT = brain_lib.cfg.vault_root
NOTES_DIR = brain_lib.cfg.vault_notes
INBOX_DIR = brain_lib.cfg.vault_inbox
MAIL_NOTE_TYPE = "interaction"
TASK_NOTE_TYPE = "task"
MAIL_CLIENT_APPLE = "apple"
MAIL_CLIENT_OUTLOOK = "outlook"
MAIL_CLIENT_AUTO = "auto"
MAIL_CLIENTS = (MAIL_CLIENT_APPLE, MAIL_CLIENT_OUTLOOK, MAIL_CLIENT_AUTO)
OUTLOOK_RUNTIME_BLOCKER = "System Events access is blocked"
MAIL_CAPTURE_VERSION = 1
MAIL_CAPTURE_SOURCE = "save-mail"
MAIL_ENRICHMENT_STATUS_PENDING = "pending"
MAIL_ENRICHMENT_VERSION = 0


def _load_mail_client(client: str) -> ModuleType:
    if client == MAIL_CLIENT_APPLE:
        return mail_applescript
    if client == MAIL_CLIENT_OUTLOOK:
        return importlib.import_module("mail_outlook_applescript")
    raise ValueError(f"Unsupported mail client: {client}")


def _is_client_running(adapter: ModuleType) -> bool:
    if hasattr(adapter, "is_mail_running"):
        return bool(adapter.is_mail_running())
    if hasattr(adapter, "is_outlook_running"):
        return bool(adapter.is_outlook_running())
    return True


def _selected_headers_for_client(client: str) -> list[dict[str, Any]]:
    adapter = _load_mail_client(client)
    return adapter.get_selected_mail_headers()


def get_selected_headers_for_client(client: str) -> tuple[str, list[dict[str, Any]]]:
    """Return selected mail headers for a concrete or auto-detected client."""
    if client != MAIL_CLIENT_AUTO:
        if client == MAIL_CLIENT_APPLE:
            return client, get_selected_mail_headers()
        return client, _selected_headers_for_client(client)

    errors: list[str] = []
    for candidate in (MAIL_CLIENT_OUTLOOK, MAIL_CLIENT_APPLE):
        adapter = _load_mail_client(candidate)
        if not _is_client_running(adapter):
            errors.append(f"{candidate}: app not running")
            continue
        try:
            return candidate, adapter.get_selected_mail_headers()
        except Exception as exc:
            if candidate == MAIL_CLIENT_OUTLOOK and OUTLOOK_RUNTIME_BLOCKER in str(exc):
                raise RuntimeError(str(exc)) from exc
            errors.append(f"{candidate}: {exc}")
    detail = "; ".join(errors) if errors else "no supported mail clients available"
    raise RuntimeError(f"No selected mail found in Outlook or Apple Mail ({detail})")


@contextmanager
def use_mail_client(client: str) -> Iterator[ModuleType]:
    """Temporarily route adapter calls used by the save flow."""
    adapter = _load_mail_client(client)
    old_values = (
        globals()["archive_mail"],
        globals()["fetch_all_mailboxes"],
        globals()["fetch_conversation_candidates"],
        globals()["fetch_mail_body"],
        globals()["get_selected_mail_headers"],
        globals()["save_attachments"],
    )
    globals()["archive_mail"] = adapter.archive_mail
    globals()["fetch_all_mailboxes"] = adapter.fetch_all_mailboxes
    globals()["fetch_mail_body"] = adapter.fetch_mail_body
    globals()["get_selected_mail_headers"] = adapter.get_selected_mail_headers
    globals()["save_attachments"] = adapter.save_attachments
    if hasattr(adapter, "fetch_conversation_candidates"):
        globals()["fetch_conversation_candidates"] = adapter.fetch_conversation_candidates
    else:
        globals()["fetch_conversation_candidates"] = lambda *args, **kwargs: []
    try:
        yield adapter
    finally:
        (
            globals()["archive_mail"],
            globals()["fetch_all_mailboxes"],
            globals()["fetch_conversation_candidates"],
            globals()["fetch_mail_body"],
            globals()["get_selected_mail_headers"],
            globals()["save_attachments"],
        ) = old_values


def mail_client_context(client: str):
    if client == MAIL_CLIENT_APPLE:
        return nullcontext()
    return use_mail_client(client)

# ── Date Parsing ─────────────────────────────────────────────────────────────


def parse_apple_date(date_str: str) -> datetime:
    """Parse AppleScript date formats."""
    cleaned = re.sub(r"^\w+,\s*", "", date_str.strip())
    cleaned = cleaned.replace(" at ", " ").replace(" om ", " ")
    for fmt in (
        "%d %B %Y %H:%M:%S",
        "%d %b %Y %H:%M:%S",
        "%d-%m-%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
    ):
        try:
            return datetime.strptime(cleaned, fmt)
        except ValueError:
            continue
    return datetime.now()


# ── Area & Project Detection ─────────────────────────────────────────────────


def detect_area(sender_email: str, to_emails: str) -> str:
    all_emails = f"{sender_email} {to_emails}"
    if "mccoy" in all_emails or "groningen.nl" in all_emails or "uu.nl" in all_emails:
        return "work"
    return "self"


def detect_project(sender_email: str, to_emails: str) -> str | None:
    match = detect_mail_project(
        {"sender_email": sender_email, "sender_display": sender_email, "to": to_emails}
    )
    if match and not match.is_ambiguous:
        return match.project
    return None


def detect_mail_project(mail: dict[str, Any]) -> MailProjectMatch | None:
    """Best-effort project detection for mail metadata."""
    try:
        return resolve_mail_project(mail)
    except Exception:
        return None


# ── Deduplication ────────────────────────────────────────────────────────────

_DEDUP_CACHE = brain_lib.ROOT / "context" / "observability" / "save-mail-dedup.json"


def _dedup_cache_read() -> dict[str, str]:
    try:
        if _DEDUP_CACHE.exists():
            cache = json.loads(_DEDUP_CACHE.read_text())
            if isinstance(cache, dict):
                return {str(k): str(v) for k, v in cache.items()}
    except Exception:
        pass
    return {}


def _dedup_cache_write(cache: dict[str, str]) -> None:
    _DEDUP_CACHE.parent.mkdir(parents=True, exist_ok=True)
    _DEDUP_CACHE.write_text(json.dumps(cache, sort_keys=True), encoding="utf-8")


def _dedup_cache_check(message_id: str, fingerprint: str = "") -> str | None:
    """Check local dedup cache (survives when SQLite indexer is down)."""
    cache = _dedup_cache_read()
    if fingerprint:
        cached = cache.get(f"fingerprint:{fingerprint}")
        if cached:
            return cached
    if message_id:
        return cache.get(message_id) or cache.get(f"message-id:{message_id.lower()}")
    return None


def _normalized_mail_addresses(value: str) -> str:
    emails = sorted({email.lower() for email in EMAIL_RE.findall(value or "")})
    if emails:
        return ",".join(emails)
    return re.sub(r"\s+", " ", (value or "").strip().lower())


def mail_dedup_fingerprint(mail: dict[str, Any]) -> str:
    """Build a stable identity for a mail when Message-ID lookup is unavailable."""
    mailbox_type = _mailbox_type(mail)
    dt = parse_apple_date(str(mail.get("date_str", "")))
    subject = normalize_subject(mail.get("subject", ""))
    to_emails = _effective_to_emails(mail, mailbox_type)
    parts = [
        "mail-dedup-v1",
        dt.strftime("%Y-%m-%d %H:%M:%S"),
        subject,
        str(mail.get("sender_email", "")).strip().lower(),
        _normalized_mail_addresses(to_emails),
        _normalized_mail_addresses(str(mail.get("cc", ""))),
    ]
    raw = "\n".join(parts)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]


def _dedup_cache_add(
    message_id: str, slug: str, mail: dict[str, Any] | None = None
) -> None:
    """Add entry to local dedup cache."""
    message_id = clean_message_id(message_id)
    fingerprint = mail_dedup_fingerprint(mail) if mail else ""
    if not message_id and not fingerprint:
        return
    try:
        cache = _dedup_cache_read()
        if message_id:
            cache[message_id] = slug
            cache[f"message-id:{message_id.lower()}"] = slug
        if fingerprint:
            cache[f"fingerprint:{fingerprint}"] = slug
        _dedup_cache_write(cache)
    except Exception:
        pass


def _sqlite_duplicate_by_message_id(message_id: str) -> str | None:
    try:
        import sqlite3

        db_path = VAULT_ROOT / ".brain-vault-index.sqlite"
        if not db_path.exists():
            return None
        conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
        try:
            conn.execute("PRAGMA query_only = ON")
            mail_link = f"message://<{message_id}>"
            mail_link_like = (
                mail_link.replace("\\", "\\\\")
                .replace("%", "\\%")
                .replace("_", "\\_")
            )
            rows = conn.execute(
                "SELECT path FROM notes WHERE type = 'interaction' AND category = 'mail' "
                "AND frontmatter_json LIKE ? ESCAPE '\\' LIMIT 1",
                (f"%{mail_link_like}%",),
            ).fetchall()
            if rows:
                from pathlib import PurePosixPath

                rows = [(PurePosixPath(rows[0][0]).stem,)]
        finally:
            conn.close()
        return rows[0][0] if rows else None
    except Exception:
        return None


def _sqlite_duplicate_by_fingerprint(fingerprint: str) -> str | None:
    if not fingerprint:
        return None
    try:
        import sqlite3

        db_path = VAULT_ROOT / ".brain-vault-index.sqlite"
        if not db_path.exists():
            return None
        conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
        try:
            conn.execute("PRAGMA query_only = ON")
            escaped = (
                fingerprint.replace("\\", "\\\\")
                .replace("%", "\\%")
                .replace("_", "\\_")
            )
            rows = conn.execute(
                "SELECT path FROM notes WHERE type = 'interaction' AND category = 'mail' "
                "AND frontmatter_json LIKE ? ESCAPE '\\' LIMIT 1",
                (f"%{escaped}%",),
            ).fetchall()
            if rows:
                from pathlib import PurePosixPath

                rows = [(PurePosixPath(rows[0][0]).stem,)]
        finally:
            conn.close()
        return rows[0][0] if rows else None
    except Exception:
        return None


def _frontmatter_matches_mail(frontmatter: dict[str, Any], mail: dict[str, Any]) -> bool:
    mailbox_type = _mailbox_type(mail)
    dt = parse_apple_date(str(mail.get("date_str", "")))
    expected_ts = dt.strftime("%Y%m%d-%H%M")
    note_ts = str(frontmatter.get("timestamp", ""))
    if note_ts != expected_ts:
        return False
    note_subject = str(
        frontmatter.get("clean_subject") or frontmatter.get("raw_subject") or ""
    )
    if normalize_subject(note_subject) != normalize_subject(mail.get("subject", "")):
        return False
    if str(frontmatter.get("from", "")).strip().lower() != str(
        mail.get("sender_email", "")
    ).strip().lower():
        return False
    note_to = _normalized_mail_addresses(str(frontmatter.get("to", "")))
    mail_to = _normalized_mail_addresses(_effective_to_emails(mail, mailbox_type))
    if note_to or mail_to:
        return note_to == mail_to
    return True


def _filesystem_duplicate_by_mail(mail: dict[str, Any], fingerprint: str) -> str | None:
    dt = parse_apple_date(str(mail.get("date_str", "")))
    ts = dt.strftime("%Y%m%d-%H%M")
    month_dir = NOTES_DIR / dt.strftime("%Y-%m")
    candidate_dirs = [month_dir, INBOX_DIR]
    candidates: list[Path] = []
    for directory in candidate_dirs:
        if directory.exists():
            candidates.extend(directory.glob(f"{ts}-mail-*.md"))

    def candidate_sort_key(path: Path) -> tuple[str, int, str]:
        base = re.sub(r"-\d+$", "", path.stem)
        has_counter = 1 if base != path.stem else 0
        return (base, has_counter, path.stem)

    for path in sorted(candidates, key=candidate_sort_key):
        try:
            frontmatter, _ = _split_note_content(path.read_text(encoding="utf-8"))
        except Exception:
            continue
        if not frontmatter:
            continue
        slug = str(frontmatter.get("slug") or path.stem)
        if fingerprint and frontmatter.get("dedup_fingerprint") == fingerprint:
            return slug
        if _frontmatter_matches_mail(frontmatter, mail):
            return slug
    return None


def is_duplicate(message_id: str, mail: dict[str, Any] | None = None) -> str | None:
    """Check if a mail already exists. Returns slug or None."""
    message_id = clean_message_id(message_id)
    fingerprint = mail_dedup_fingerprint(mail) if mail else ""
    if not message_id and not fingerprint:
        return None

    cached = _dedup_cache_check(message_id, fingerprint)
    if cached:
        return cached

    existing = _sqlite_duplicate_by_message_id(message_id) if message_id else None
    if not existing and fingerprint:
        existing = _sqlite_duplicate_by_fingerprint(fingerprint)
    if not existing and mail:
        existing = _filesystem_duplicate_by_mail(mail, fingerprint)
    if existing:
        _dedup_cache_add(message_id, existing, mail=mail)
    return existing


def find_thread_notes(entity_slug: str, subject: str) -> list[str]:
    """Find existing mail notes in the same thread (same entity + cleaned subject)."""
    cleaned = re.sub(
        r"^((Re|Fwd|FW|AW|Antw|SV):\s*)+", "", subject, flags=re.IGNORECASE
    ).strip()
    if not cleaned or len(cleaned) < 3:
        return []
    try:
        import sqlite3

        db_path = VAULT_ROOT / ".brain-vault-index.sqlite"
        if not db_path.exists():
            return []
        conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
        try:
            conn.execute("PRAGMA query_only = ON")
            escaped = (
                cleaned.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
            )
            entity_json = f'["{entity_slug}"]'
            rows = conn.execute(
                "SELECT slug FROM notes WHERE type = 'interaction' AND category = 'mail' "
                "AND entity = ? "
                "AND title LIKE ? ESCAPE '\\' ORDER BY created DESC LIMIT 10",
                (entity_json, f"%{escaped}%"),
            ).fetchall()
        finally:
            conn.close()
        return [r[0] for r in rows]
    except Exception:
        return []


# ── Conversation Detection ──────────────────────────────────────────────────

REPLY_PREFIX_RE = re.compile(
    r"^((Re|Fwd|FW|AW|Antw|SV):\s*)+", flags=re.IGNORECASE
)
CONVERSATION_REPLY_PREFIX_RE = re.compile(
    r"^((Re|AW|Antw|SV):\s*)+", flags=re.IGNORECASE
)
MESSAGE_ID_RE = re.compile(r"<([^<>\s]+)>")
EMAIL_RE = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)
GENERIC_SUBJECTS = {
    "hi",
    "hoi",
    "hello",
    "vraag",
    "question",
    "meeting",
    "update",
    "invoice",
    "factuur",
}
_SELF_EMAIL_SET = {email.lower() for email in SELF_EMAILS}


def strip_reply_prefixes(subject: str) -> str:
    return REPLY_PREFIX_RE.sub("", subject or "").strip()


def has_reply_prefix(subject: str) -> bool:
    return bool(CONVERSATION_REPLY_PREFIX_RE.match(subject or ""))


def normalize_subject(subject: str) -> str:
    cleaned = strip_reply_prefixes(subject)
    cleaned = re.sub(r"\s+", " ", cleaned).strip().lower()
    return cleaned


def clean_mail_subject(subject: str) -> str:
    """Return a conservative display/search subject while preserving raw_subject."""
    cleaned = strip_reply_prefixes(subject or "")
    cleaned = re.sub(
        r"^\s*(\[(external|extern|ext|spam|bulk)\]\s*)+",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned or (subject or "Geen onderwerp").strip() or "Geen onderwerp"


def is_generic_subject(subject: str) -> bool:
    normalized = normalize_subject(subject)
    return len(normalized) < 5 or normalized in GENERIC_SUBJECTS


def parse_mail_headers(raw_headers: str) -> dict[str, str]:
    """Parse unfolded RFC-style mail headers into a lowercase dict."""
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
        normalized_key = key.strip().lower()
        normalized_value = value.strip()
        if normalized_key in headers:
            headers[normalized_key] = headers[normalized_key] + "\n" + normalized_value
        else:
            headers[normalized_key] = normalized_value
    return headers


def clean_message_id(value: str) -> str:
    cleaned = (value or "").strip().strip("<>").strip()
    if cleaned.startswith("message://"):
        cleaned = cleaned.removeprefix("message://").strip("<>")
    return cleaned


def _has_reply_context(mail: dict[str, Any]) -> bool:
    headers = parse_mail_headers(mail.get("all_headers", ""))
    return has_reply_prefix(mail.get("subject", "")) or bool(headers.get("in-reply-to"))


def existing_thread_slug_map(message_ids: list[str]) -> dict[str, str]:
    """Return already-saved thread slugs keyed by normalized message id."""
    existing: dict[str, str] = {}
    for message_id in message_ids:
        cleaned = clean_message_id(message_id)
        key = cleaned.lower()
        if not cleaned or key in existing:
            continue
        slug = is_duplicate(cleaned)
        if slug:
            existing[key] = slug
    return existing


def extract_message_ids(value: str) -> list[str]:
    if not value:
        return []
    matches = MESSAGE_ID_RE.findall(value)
    if not matches:
        matches = [token for token in re.split(r"[\s,]+", value) if "@" in token]

    seen: set[str] = set()
    ids: list[str] = []
    for match in matches:
        cleaned = clean_message_id(match)
        key = cleaned.lower()
        if cleaned and key not in seen:
            seen.add(key)
            ids.append(cleaned)
    return ids


def conversation_message_ids(mail: dict[str, Any]) -> list[str]:
    headers = parse_mail_headers(mail.get("all_headers", ""))
    values = [
        mail.get("message_id", ""),
        headers.get("message-id", ""),
        headers.get("in-reply-to", ""),
        headers.get("references", ""),
    ]

    seen: set[str] = set()
    ids: list[str] = []
    for value in values:
        candidates = extract_message_ids(value)
        if not candidates and value:
            candidates = [clean_message_id(value)]
        for message_id in candidates:
            key = message_id.lower()
            if message_id and key not in seen:
                seen.add(key)
                ids.append(message_id)
    return ids


def _thread_index_root(value: str) -> str:
    compact = re.sub(r"\s+", "", value or "")
    if not compact:
        return ""
    return compact[:44] if len(compact) >= 44 else compact


def mail_thread_metadata(mail: dict[str, Any]) -> dict[str, Any]:
    """Return stable thread metadata when headers or safe fallback provide evidence."""
    headers = parse_mail_headers(mail.get("all_headers", ""))
    references = extract_message_ids(headers.get("references", ""))
    in_reply_to = extract_message_ids(headers.get("in-reply-to", ""))
    thread_topic = headers.get("thread-topic", "")
    thread_index_root = _thread_index_root(headers.get("thread-index", ""))
    message_ids = conversation_message_ids(mail)

    metadata: dict[str, Any] = {
        "references": references,
        "in_reply_to": in_reply_to[0] if in_reply_to else "",
        "root_message_id": "",
        "thread_id": "",
        "thread_source": "",
        "thread_topic": normalize_subject(thread_topic) if thread_topic else "",
        "thread_index_root": thread_index_root,
    }

    if references:
        root = references[0]
        metadata.update(
            {
                "root_message_id": root,
                "thread_id": f"message-id:{root.lower()}",
                "thread_source": "references",
            }
        )
        return metadata

    if in_reply_to:
        root = in_reply_to[0]
        metadata.update(
            {
                "root_message_id": root,
                "thread_id": f"message-id:{root.lower()}",
                "thread_source": "in-reply-to",
            }
        )
        return metadata

    if thread_index_root:
        metadata.update(
            {
                "root_message_id": message_ids[-1] if message_ids else "",
                "thread_id": f"thread-index:{thread_index_root.lower()}",
                "thread_source": "thread-index",
            }
        )
        return metadata

    subject = normalize_subject(mail.get("subject", ""))
    participants = sorted(_external_participants(mail))
    if has_reply_prefix(mail.get("subject", "")) and subject and participants:
        if not is_generic_subject(subject):
            raw_key = "|".join([subject, *participants])
            digest = hashlib.sha1(raw_key.encode("utf-8")).hexdigest()[:16]
            metadata.update(
                {
                    "root_message_id": message_ids[-1] if message_ids else "",
                    "thread_id": f"subject-participants:{digest}",
                    "thread_source": "reply-subject-participants",
                }
            )
    return metadata


def _is_self_email(email: str) -> bool:
    normalized = email.lower()
    if normalized in _SELF_EMAIL_SET:
        return True
    for pattern in _SELF_EMAIL_SET:
        if "*" not in pattern:
            continue
        regex = "^" + re.escape(pattern).replace("\\*", ".*") + "$"
        if re.match(regex, normalized):
            return True
    return False


def mail_participants(mail: dict[str, Any]) -> set[str]:
    headers = parse_mail_headers(mail.get("all_headers", ""))
    fields = [
        mail.get("sender_email", ""),
        mail.get("sender_display", ""),
        mail.get("to", ""),
        mail.get("cc", ""),
        headers.get("from", ""),
        headers.get("to", ""),
        headers.get("cc", ""),
        headers.get("reply-to", ""),
    ]
    return {email.lower() for email in EMAIL_RE.findall(" ".join(fields))}


def _external_participants(mail: dict[str, Any]) -> set[str]:
    return {email for email in mail_participants(mail) if not _is_self_email(email)}


def _participants_overlap(selected: dict[str, Any], candidate: dict[str, Any]) -> bool:
    selected_external = _external_participants(selected)
    candidate_external = _external_participants(candidate)
    if selected_external and candidate_external:
        return bool(selected_external & candidate_external)
    return False


def _conversation_signature(mail: dict[str, Any]) -> dict[str, Any]:
    headers = parse_mail_headers(mail.get("all_headers", ""))
    subject = mail.get("subject", "")
    thread_topic = headers.get("thread-topic", "")
    return {
        "primary_id": clean_message_id(mail.get("message_id", "")).lower(),
        "message_ids": {mid.lower() for mid in conversation_message_ids(mail)},
        "thread_index_root": _thread_index_root(headers.get("thread-index", "")),
        "thread_topic": normalize_subject(thread_topic) if thread_topic else "",
        "subject": normalize_subject(subject),
        "has_reply_prefix": has_reply_prefix(subject),
    }


def _conversation_match_reason(
    selected_mail: dict[str, Any],
    candidate_mail: dict[str, Any],
    *,
    force_subject: bool = False,
) -> str | None:
    selected = _conversation_signature(selected_mail)
    candidate = _conversation_signature(candidate_mail)
    if candidate["primary_id"] and candidate["primary_id"] == selected["primary_id"]:
        return "selected"
    if selected["message_ids"] and candidate["message_ids"]:
        if selected["message_ids"] & candidate["message_ids"]:
            return "message-id"
    if (
        selected["thread_index_root"]
        and candidate["thread_index_root"]
        and selected["thread_index_root"] == candidate["thread_index_root"]
    ):
        return "thread-index"
    if (
        selected["thread_topic"]
        and candidate["thread_topic"]
        and selected["thread_topic"] == candidate["thread_topic"]
        and _participants_overlap(selected_mail, candidate_mail)
    ):
        return "thread-topic"
    subject_matches = (
        selected["subject"]
        and candidate["subject"]
        and selected["subject"] == candidate["subject"]
        and not is_generic_subject(selected["subject"])
        and _participants_overlap(selected_mail, candidate_mail)
    )
    if subject_matches and (
        force_subject or selected["has_reply_prefix"] or candidate["has_reply_prefix"]
    ):
        return "reply-subject"
    return None


def conversation_fetch_hints(
    selected_mail: dict[str, Any], *, force_subject: bool = False
) -> dict[str, Any]:
    headers = parse_mail_headers(selected_mail.get("all_headers", ""))
    thread_topic = headers.get("thread-topic", "")
    subject_hint = thread_topic or strip_reply_prefixes(selected_mail.get("subject", ""))
    include_subject = bool(force_subject or _has_reply_context(selected_mail))
    if is_generic_subject(subject_hint):
        include_subject = force_subject
    return {
        "message_ids": list(conversation_message_ids(selected_mail)),
        "subject_hint": strip_reply_prefixes(subject_hint),
        "include_subject": include_subject,
    }


def conversation_lookup_ids(
    selected_mail: dict[str, Any], *, force: bool = False
) -> list[str]:
    """Return message ids worth looking up outside the selected mail itself."""
    if not force and not _has_reply_context(selected_mail):
        return []

    selected_id = clean_message_id(selected_mail.get("message_id", "")).lower()
    ids: list[str] = []
    seen: set[str] = set()
    for message_id in conversation_message_ids(selected_mail):
        cleaned = clean_message_id(message_id)
        key = cleaned.lower()
        if not cleaned or key == selected_id or key in seen:
            continue
        seen.add(key)
        ids.append(cleaned)
    return ids


def select_conversation_mails(
    selected_mail: dict[str, Any],
    candidates: list[dict[str, Any]],
    *,
    force: bool = False,
) -> dict[str, Any]:
    by_id: dict[str, dict[str, Any]] = {}
    selected_id = clean_message_id(selected_mail.get("message_id", "")).lower()
    if selected_id:
        by_id[selected_id] = selected_mail
    for candidate in candidates:
        candidate_id = clean_message_id(candidate.get("message_id", "")).lower()
        if candidate_id and candidate_id not in by_id:
            by_id[candidate_id] = candidate

    matches: list[dict[str, Any]] = []
    reasons: set[str] = set()
    for candidate in by_id.values():
        reason = _conversation_match_reason(
            selected_mail, candidate, force_subject=force
        )
        if reason:
            matches.append(candidate)
            if reason != "selected":
                reasons.add(reason)

    if len(matches) < 2 or (not force and not reasons):
        return {
            "is_conversation": False,
            "reason": "single",
            "mails": [selected_mail],
            "reasons": [],
        }

    return {
        "is_conversation": True,
        "reason": ", ".join(sorted(reasons)) or "forced",
        "mails": sorted(matches, key=lambda mail: parse_apple_date(mail["date_str"])),
        "reasons": sorted(reasons),
    }


# ── Note Creation ────────────────────────────────────────────────────────────


def _subject_words(subject: str, max_words: int = 3) -> str:
    """Extract first N meaningful words from subject for slug context."""
    cleaned = strip_reply_prefixes(subject)
    cleaned = re.sub(r"\[.*?\]", "", cleaned).strip()
    words = re.findall(r"[a-z0-9]+", cleaned.lower())
    words = [w for w in words if len(w) >= 2]
    return "-".join(words[:max_words])


def make_slug(dt: datetime, entity_slug: str, subject: str) -> tuple[str, str]:
    """Create unique slug: YYYYMMDD-HHMM-mail-entity-word1-word2-word3."""
    ts = dt.strftime("%Y%m%d-%H%M")
    ctx = _subject_words(subject)
    if ctx:
        base = f"{ts}-mail-{entity_slug}-{ctx}"
    else:
        base = f"{ts}-mail-{entity_slug}"

    note_dir = brain_lib.canonical_note_dir(MAIL_NOTE_TYPE, ts, create=True)

    filepath = note_dir / f"{base}.md"
    if not filepath.exists():
        return base, ts
    counter = 2
    while (note_dir / f"{base}-{counter}.md").exists():
        counter += 1
    return f"{base}-{counter}", ts


def infer_mail_direction(mail: dict[str, Any]) -> str:
    """Return the same sent/received label basis used by create_mail_note."""
    try:
        _, direction = resolve_entity(mail.get("sender_email", ""), mail.get("to", ""))
        return direction
    except Exception:
        return "sent" if _is_self_email(mail.get("sender_email", "")) else "received"


def thread_timeline_label(direction: str) -> str:
    if direction == "sent":
        return "Sent"
    if direction == "received":
        return "Received"
    return "Mail"


def _sort_thread_entries(entries: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(entries, key=lambda entry: str(entry.get("slug", ""))[:13])


def render_thread_timeline(
    entries: list[dict[str, Any]], current_slug: str | None = None
) -> list[str]:
    """Render a complete thread timeline block."""
    unique_entries: list[dict[str, Any]] = []
    seen: set[str] = set()
    for entry in _sort_thread_entries(entries):
        slug = str(entry.get("slug", "")).strip()
        if not slug or slug in seen:
            continue
        seen.add(slug)
        unique_entries.append(entry)

    if len(unique_entries) < 2:
        return []

    lines = ["🔗 Thread:"]
    for entry in unique_entries:
        slug = str(entry["slug"])
        label = thread_timeline_label(str(entry.get("direction", "")))
        suffix = " (current)" if current_slug and slug == current_slug else ""
        lines.append(f"- {label} [[{slug}]]{suffix}")
    return lines


def find_note_path_by_slug(slug: str) -> Path | None:
    inbox_path = INBOX_DIR / f"{slug}.md"
    if inbox_path.exists():
        return inbox_path
    for candidate in NOTES_DIR.rglob(f"{slug}.md"):
        return candidate

    try:
        import sqlite3

        db_path = VAULT_ROOT / ".brain-vault-index.sqlite"
        if not db_path.exists():
            return None
        conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
        try:
            conn.execute("PRAGMA query_only = ON")
            row = conn.execute(
                "SELECT path FROM notes WHERE slug = ? LIMIT 1",
                (slug,),
            ).fetchone()
        finally:
            conn.close()
        if not row:
            return None
        candidate = Path(row[0])
        if not candidate.is_absolute():
            candidate = VAULT_ROOT / candidate
        return candidate if candidate.exists() else None
    except Exception:
        return None


def _split_note_content(content: str) -> tuple[dict[str, Any], str]:
    match = re.match(r"^---\n(.*?)\n---\n?(.*)$", content, flags=re.DOTALL)
    if not match:
        return {}, content
    raw_fm, body = match.groups()
    frontmatter = yaml.safe_load(raw_fm) or {}
    if not isinstance(frontmatter, dict):
        frontmatter = {}
    return frontmatter, body


def _thread_entry_from_slug(slug: str) -> dict[str, Any]:
    path = find_note_path_by_slug(slug)
    direction = ""
    if path and path.exists():
        try:
            fm, _ = _split_note_content(path.read_text(encoding="utf-8"))
            direction = str(fm.get("direction", ""))
        except Exception:
            direction = ""
    return {"slug": slug, "direction": direction, "path": str(path) if path else ""}


def replace_thread_timeline_block(
    body: str,
    entries: list[dict[str, Any]],
    *,
    current_slug: str,
) -> str:
    """Replace legacy or generated thread blocks without touching the mail body."""
    timeline_lines = render_thread_timeline(entries, current_slug)
    lines = body.split("\n")

    stripped: list[str] = []
    i = 0
    while i < len(lines):
        if lines[i].startswith("🔗 Thread:"):
            i += 1
            while i < len(lines) and lines[i].startswith("- "):
                i += 1
            continue
        stripped.append(lines[i])
        i += 1

    if not timeline_lines:
        return "\n".join(stripped)

    insert_at = None
    for idx, line in enumerate(stripped):
        if line.startswith("[📩 Open in Mail]"):
            insert_at = idx + 1
            while insert_at < len(stripped) and stripped[insert_at].startswith(
                "📋 Task:"
            ):
                insert_at += 1
            break
    if insert_at is None:
        for idx, line in enumerate(stripped):
            if line == "---":
                insert_at = idx
                break
    if insert_at is None:
        insert_at = len(stripped)

    stripped[insert_at:insert_at] = timeline_lines
    return "\n".join(stripped)


def update_note_thread_timeline(
    path: Path,
    entries: list[dict[str, Any]],
    *,
    current_slug: str,
) -> bool:
    """Update one existing mail note with a complete thread timeline."""
    if not path.exists():
        return False
    content = path.read_text(encoding="utf-8")
    frontmatter, body = _split_note_content(content)
    if not frontmatter:
        return False

    slugs = [str(entry.get("slug", "")) for entry in _sort_thread_entries(entries)]
    thread_slugs = [slug for slug in slugs if slug and slug != current_slug]
    if thread_slugs:
        frontmatter["thread"] = thread_slugs
    else:
        frontmatter.pop("thread", None)

    body = replace_thread_timeline_block(body, entries, current_slug=current_slug)
    vnw.write_vault_note(path, frontmatter, body)
    return True


def update_conversation_thread_timeline(entries: list[dict[str, Any]]) -> int:
    """Update all notes in a conversation with the same complete timeline."""
    if len(entries) < 2:
        return 0

    updated = 0
    for entry in _sort_thread_entries(entries):
        slug = str(entry.get("slug", ""))
        if not slug:
            continue
        path_value = entry.get("path")
        path = Path(path_value) if path_value else find_note_path_by_slug(slug)
        if not path:
            continue
        if update_note_thread_timeline(path, entries, current_slug=slug):
            updated += 1
    return updated


def create_mail_note(
    mail: dict[str, str],
    create_task: bool = False,
    extra_thread_slugs: list[str] | None = None,
) -> dict:
    """Create vault interaction note. Returns result dict."""
    dt = parse_apple_date(mail["date_str"])
    sender_email = mail["sender_email"]
    mailbox_type = _mailbox_type(mail)
    to_emails = _effective_to_emails(mail, mailbox_type)
    subject = mail.get("subject", "Geen onderwerp")
    clean_subject = clean_mail_subject(subject)
    message_id = mail["message_id"]
    mail_client = mail.get("mail_client", "apple")
    calendar_invite = is_calendar_invite(mail)
    sender_domain = _email_domain(sender_email)

    entity_resolution = resolve_entity_details(sender_email, to_emails)
    entity_slug = entity_resolution.slug
    direction = entity_resolution.direction
    entity_source = entity_resolution.source
    entity_confidence = entity_resolution.confidence
    if entity_slug == "unknown-sender" and "@" not in sender_email:
        display_slug = _entity_slug_from_display_name(
            mail.get("sender_display", "") or sender_email
        )
        if display_slug:
            entity_slug = display_slug
            entity_source = "outlook-display-name"
            entity_confidence = 0.65
            _ensure_display_entity_note(
                entity_slug,
                mail.get("sender_display", "") or sender_email,
                area=_area_from_source_account(mail.get("account", "")),
            )
    project_match = detect_mail_project(mail)
    area = (
        project_match.area
        if project_match and project_match.area
        else detect_area(sender_email, to_emails)
    )
    project = (
        project_match.project
        if project_match and not project_match.is_ambiguous
        else None
    )
    project_slug = project
    project_code = (
        brain_lib.resolve_project_code(project_slug) if project_slug else None
    )

    thread_slugs = find_thread_notes(entity_slug, subject)
    for thread_slug in extra_thread_slugs or []:
        if thread_slug not in thread_slugs:
            thread_slugs.append(thread_slug)

    slug, ts = make_slug(dt, entity_slug, clean_subject)
    filepath = brain_lib.canonical_note_path(
        MAIL_NOTE_TYPE, slug, ts, create_parent=True
    )

    thread_slugs = [s for s in thread_slugs if s != slug]

    created = dt.strftime("%Y-%m-%d")
    date_display = dt.strftime("%-d %b %Y at %H:%M")
    emoji = "📤" if direction == "sent" else "📥"
    escaped_subject = escape_title(subject)
    escaped_clean_subject = escape_title(clean_subject)

    # Save attachments
    att_links: list[str] = []
    att_count = int(mail.get("att_count", "0"))
    if att_count > 0 and mail_client != MAIL_CLIENT_OUTLOOK:
        att_links = save_attachments(
            message_id,
            ts,
            account=mail.get("account", ""),
            mailbox=mail.get("mailbox", ""),
            filename_prefix=slug,
        )

    # Build frontmatter
    topics, topics_source, topics_confidence = _topic_routing_metadata(
        entity_slug,
        subject,
        calendar_invite=calendar_invite,
    )
    thread_metadata = mail_thread_metadata(mail)
    dedup_fingerprint = mail_dedup_fingerprint(mail)
    fm: dict[str, Any] = {
        "type": "interaction",
        "category": "mail",
        "created": created,
        "slug": slug,
        "timestamp": ts,
        "area": area,
        "capture_source": MAIL_CAPTURE_SOURCE,
        "capture_version": MAIL_CAPTURE_VERSION,
        "clean_subject": escaped_clean_subject,
        "direction": direction,
        "dedup_fingerprint": dedup_fingerprint,
        "enrichment_status": MAIL_ENRICHMENT_STATUS_PENDING,
        "enrichment_version": MAIL_ENRICHMENT_VERSION,
        "entity": [entity_slug],
        "entity_confidence": entity_confidence,
        "entity_source": entity_source,
        "from": sender_email,
        "mail_link": f"message://<{message_id}>",
        "mail_client": mail_client,
        "mailbox_type": mailbox_type,
        "raw_subject": escaped_subject,
        "sender_domain": sender_domain,
        "source_account": mail.get("account", ""),
        "source_mailbox": mail.get("mailbox", ""),
        "title": escaped_clean_subject,
        "to": to_emails,
        "topics": topics,
        "topics_confidence": topics_confidence,
        "topics_source": topics_source,
    }
    if calendar_invite:
        fm["calendar_invite"] = True
    if project_slug:
        fm["project"] = project_code
        fm["project_slug"] = project_slug
        if project_match:
            fm["project_source"] = project_match.source
            fm["project_confidence"] = project_match.confidence
            fm["project_match"] = project_match.matched
    if mail.get("cc"):
        fm["cc"] = mail["cc"]
    if att_links:
        att_names = []
        for link in att_links:
            att_name = link.split("|")[-1].rstrip("]]") if "|" in link else link
            att_names.append(att_name)
        fm["attachments"] = att_names
    if thread_slugs:
        fm["thread"] = thread_slugs
    if thread_metadata.get("thread_id"):
        fm["thread_id"] = thread_metadata["thread_id"]
        fm["thread_source"] = thread_metadata["thread_source"]
        if thread_metadata.get("root_message_id"):
            fm["root_message_id"] = thread_metadata["root_message_id"]
        if thread_metadata.get("in_reply_to"):
            fm["in_reply_to"] = thread_metadata["in_reply_to"]
        if thread_metadata.get("references"):
            fm["references"] = thread_metadata["references"]
        if thread_metadata.get("thread_topic"):
            fm["thread_topic"] = thread_metadata["thread_topic"]
        if thread_metadata.get("thread_index_root"):
            fm["thread_index_root"] = thread_metadata["thread_index_root"]

    # Build body
    if direction == "sent":
        from_line = f"📧 From: Me ({sender_email})"
        to_line = f"📬 To: [[{entity_slug}]] ({to_emails})"
    else:
        from_line = f"📧 From: [[{entity_slug}]] ({sender_email})"
        to_line = f"📬 To: {to_emails}"

    mail_link_label = "Open in Outlook" if mail_client == MAIL_CLIENT_OUTLOOK else "Open in Mail"

    body_lines = [
        "",
            f"# {emoji} {clean_subject}",
        "",
        from_line,
        to_line,
    ]
    if mail.get("cc"):
        body_lines.append(f"📋 CC: {mail['cc']}")
    if calendar_invite:
        body_lines.append(f"🗓️ Calendar invite ({mailbox_type})")
    thread_body: list[str] = []
    if thread_slugs:
        thread_entries = [_thread_entry_from_slug(thread_slug) for thread_slug in thread_slugs]
        thread_entries.append({"slug": slug, "direction": direction, "path": str(filepath)})
        thread_body.extend(render_thread_timeline(thread_entries, slug))
    body_lines.extend(
        [
            f"📅 {date_display}",
            f"[📩 {mail_link_label}](message://<{message_id}>)",
            *thread_body,
            "",
            "---",
            "",
            mail.get("body", ""),
        ]
    )

    # Extract links from body
    links = re.findall(r'https?://[^\s<>")\]]+', mail.get("body", ""))
    if links:
        body_lines.extend(["", "## 🔗 Links in Mail", ""])
        seen = set()
        for link in links:
            if link not in seen:
                body_lines.append(f"- <{link}>")
                seen.add(link)

    # Attachment section
    if att_links:
        body_lines.extend(["", "## 📎 Bijlagen", ""])
        for link in att_links:
            if re.search(r"\.(png|jpg|jpeg|gif|webp|svg)\|", link, re.IGNORECASE):
                body_lines.append(f"!{link}")
            else:
                body_lines.append(f"- {link}")

    body = "\n".join(body_lines)
    vnw.write_vault_note(filepath, fm, body)

    now_ts = datetime.now().strftime("%Y%m%d-%H%M")
    brain_lib.link_note_in_daily(slug, timestamp=now_ts)

    _dedup_cache_add(message_id, slug, mail=mail)

    result = {
        "slug": slug,
        "entity": entity_slug,
        "entity_source": entity_source,
        "entity_confidence": entity_confidence,
        "direction": direction,
        "area": area,
        "project": project_code,
        "project_slug": project_slug,
        "topics": topics,
        "topics_source": topics_source,
        "topics_confidence": topics_confidence,
        "thread": thread_slugs,
        "thread_id": thread_metadata.get("thread_id", ""),
        "thread_source": thread_metadata.get("thread_source", ""),
        "root_message_id": thread_metadata.get("root_message_id", ""),
        "in_reply_to": thread_metadata.get("in_reply_to", ""),
        "references": thread_metadata.get("references", []),
        "thread_topic": thread_metadata.get("thread_topic", ""),
        "thread_index_root": thread_metadata.get("thread_index_root", ""),
        "attachments": len(att_links),
        "subject": subject,
        "raw_subject": subject,
        "clean_subject": clean_subject,
        "sender_domain": sender_domain,
        "capture_source": MAIL_CAPTURE_SOURCE,
        "capture_version": MAIL_CAPTURE_VERSION,
        "enrichment_status": MAIL_ENRICHMENT_STATUS_PENDING,
        "enrichment_version": MAIL_ENRICHMENT_VERSION,
        "path": str(filepath),
        "obsidian_file": brain_lib.obsidian_file_ref(filepath),
    }
    if project_match:
        result["project_source"] = project_match.source
        result["project_confidence"] = project_match.confidence
        result["project_match"] = project_match.matched
        if project_match.candidates:
            result["project_candidates"] = list(project_match.candidates)

    if create_task:
        task_result = create_follow_up_task(
            subject=subject,
            entity_slug=entity_slug,
            mail_slug=slug,
            area=area,
            project=project_code,
        )
        result["task_slug"] = task_result["slug"]
        result["task_path"] = task_result["path"]
        _append_task_link(filepath, task_result["slug"])
        brain_lib.link_note_in_daily(task_result["slug"])

    note_bytes = _file_size(filepath)
    if note_bytes is not None:
        result["note_bytes"] = note_bytes
        result["note_size"] = format_data_size(note_bytes)

    return result


def create_follow_up_task(
    subject: str,
    entity_slug: str,
    mail_slug: str,
    area: str,
    project: str | None,
) -> dict:
    """Create a follow-up task note linked to the mail."""
    now = datetime.now()
    created = now.strftime("%Y-%m-%d")

    clean_subject = re.sub(
        r"^(Re:|Fwd:|FW:|AW:|Antw:|SV:)\s*", "", subject, flags=re.IGNORECASE
    ).strip()
    task_title = f"Follow-up: {clean_subject}"

    note_dir = brain_lib.canonical_note_dir(TASK_NOTE_TYPE, now, create=True)
    task_slug, ts = vnw.make_slug(
        now, TASK_NOTE_TYPE, task_title, notes_dir=note_dir, max_len=50
    )
    filepath = brain_lib.canonical_note_path(
        TASK_NOTE_TYPE, task_slug, ts, create_parent=True
    )

    fm: dict[str, Any] = {
        "type": "task",
        "category": "screen",
        "created": created,
        "slug": task_slug,
        "timestamp": ts,
        "status": "to-do",
        "due": created,
        "area": area,
        "entity": [entity_slug],
        "title": escape_title(task_title),
    }
    if project:
        fm["project"] = project

    status_header = brain_lib.format_status_header("\U0001f534 to-do", now)
    source_quote = brain_lib.format_source_quote("Apple Mail", now)
    body_lines = [
        f"# {task_title}",
        "",
        f"📧 [[{mail_slug}]]",
        f"👤 [[{entity_slug}]]",
        "",
        status_header,
        "",
        "---",
        source_quote,
        "",
    ]

    body = "\n".join(body_lines)
    vnw.write_vault_note(filepath, fm, body)

    return {
        "slug": task_slug,
        "path": str(filepath),
        "obsidian_file": brain_lib.obsidian_file_ref(filepath),
    }


def _append_task_link(mail_path: Path, task_slug: str) -> None:
    """Append task reference to the mail note body."""
    content = mail_path.read_text(encoding="utf-8")
    marker = "[📩 Open in "
    if marker in content:
        lines = content.split("\n")
        for i, line in enumerate(lines):
            if marker in line:
                lines.insert(i + 1, f"📋 Task: [[{task_slug}]]")
                break
        content = "\n".join(lines)
        mail_path.write_text(content, encoding="utf-8")


# ── Calendar Invite Detection ───────────────────────────────────────────────

CALENDAR_SUBJECT_PREFIXES = (
    "Invitation:",
    "Accepted:",
    "Declined:",
    "Tentative:",
    "Updated:",
    "Cancelled:",
    "Canceled:",
)
CALENDAR_SENDERS = (
    "calendar-notification",
    "calendar.google.com",
    "calendar@outlook.com",
    "noreply@calendar.proton.me",
)


def is_calendar_invite(mail: dict) -> bool:
    """Check if a mail is a calendar invite/response.

    Detects four patterns:
    1. Standard subject prefixes (Invitation:, Accepted:, etc.)
    2. Known calendar system senders (calendar.google.com, etc.)
    3. .ics attachment (Outlook/Exchange direct invites from regular senders)
    4. has_calendar marker when available
    """
    subject = mail.get("subject", "")
    sender = mail.get("sender_email", "")
    att_names = mail.get("att_names", "")
    return (
        any(subject.startswith(p) for p in CALENDAR_SUBJECT_PREFIXES)
        or any(s in sender for s in CALENDAR_SENDERS)
        or ".ics" in att_names.lower()
        or mail.get("has_calendar") is True
    )


def should_create_follow_up_task(
    mail: dict[str, Any],
    *,
    force_task: bool = False,
    mailbox_type: str | None = None,
) -> bool:
    """Return true when a saved mail should get a follow-up task."""
    if force_task:
        return True
    effective_mailbox = mailbox_type or mail.get("mailbox_type") or "inbox"
    return bool(mail.get("is_flagged")) and effective_mailbox == "inbox"


def _mailbox_type(mail: dict[str, Any]) -> str:
    mailbox_type = mail.get("mailbox_type")
    if mailbox_type:
        return str(mailbox_type)
    return classify_mailbox_type(mail.get("mailbox", ""), mail.get("account", ""))


def _entity_slug_from_display_name(value: str) -> str | None:
    """Derive a conservative entity slug when Outlook only exposes a display name."""
    text = re.sub(
        r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}",
        " ",
        value or "",
        flags=re.I,
    )
    text = re.sub(r"\+\d+\s+others?", " ", text, flags=re.I)
    parts = re.findall(r"[a-z0-9]+", text.lower())
    parts = [part for part in parts if part not in {"unknown", "sender"}]
    if not parts:
        return None
    return "-".join(parts[:6])


def _looks_like_email(value: str) -> bool:
    return bool(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", (value or "").strip()))


def _email_domain(value: str) -> str:
    value = (value or "").strip().lower()
    if "@" not in value:
        return ""
    return value.rsplit("@", 1)[-1]


def _topic_routing_metadata(
    entity_slug: str,
    subject: str,
    *,
    calendar_invite: bool,
) -> tuple[list[str], str, float]:
    topics = set(suggest_topics(entity_slug, subject))
    sources: list[str] = []
    confidences: list[float] = []
    if topics:
        sources.append("entity-or-subject")
        confidences.append(0.7)
    if calendar_invite:
        topics.add("calendar")
        sources.append("calendar-invite")
        confidences.append(1.0)
    if not topics:
        return [], "none", 0.0
    return sorted(topics), ",".join(sources), max(confidences)


def _effective_to_emails(mail: dict[str, Any], mailbox_type: str) -> str:
    to_emails = str(mail.get("to", "") or "").strip()
    if to_emails:
        return to_emails
    account = str(mail.get("account", "") or "").strip().lower()
    if mailbox_type == "inbox" and _looks_like_email(account):
        return account
    return ""


def _area_from_source_account(account: str) -> str:
    account_lower = (account or "").strip().lower()
    if account_lower.endswith("@bam.com") or "mccoy" in account_lower:
        return "work"
    return "self"


def _display_entity_title(value: str, slug: str) -> str:
    raw = re.sub(r"\s+", " ", value or "").strip()
    if not raw:
        raw = slug.replace("-", " ")
    words = []
    known = {"askit": "AskIT", "servicenow": "ServiceNow", "bam": "BAM"}
    for word in raw.split():
        words.append(known.get(word.lower(), word[:1].upper() + word[1:]))
    return " ".join(words)


def _ensure_display_entity_note(entity_slug: str, display_name: str, *, area: str) -> None:
    """Create a minimal entity for display-only Outlook senders."""
    entity_dir = brain_lib.cfg.vault_entities
    entity_path = entity_dir / f"{entity_slug}.md"
    if entity_path.exists():
        return

    now = datetime.now()
    title = _display_entity_title(display_name, entity_slug)
    fm = {
        "type": "entity",
        "category": "company",
        "created": now.strftime("%Y-%m-%d"),
        "slug": entity_slug,
        "timestamp": now.strftime("%Y%m%d-%H%M"),
        "area": area,
        "source": "auto-created-from-outlook-display-name",
        "title": escape_title(title),
    }
    body = f"# {title}\n"
    vnw.write_vault_note(entity_path, fm, body)


def _should_archive_after_save(
    mail: dict[str, Any], *, no_archive: bool, single_mode: bool
) -> bool:
    if no_archive:
        return False
    if not clean_message_id(mail.get("message_id", "")):
        return False
    return _mailbox_type(mail) == "inbox"


def process_mail_record(
    mail: dict[str, Any],
    *,
    force_task: bool = False,
    no_archive: bool = False,
    single_mode: bool = False,
    extra_thread_slugs: list[str] | None = None,
    verbose: bool = True,
    prefix: str = "",
) -> dict[str, Any]:
    """Save one mail record and optionally archive it."""
    mail["mailbox_type"] = _mailbox_type(mail)
    subject = mail.get("subject", "?")
    message_id = mail["message_id"]

    existing = is_duplicate(message_id, mail=mail)
    _timed("dedup")
    if existing:
        now_ts = datetime.now().strftime("%Y%m%d-%H%M")
        if single_mode:
            brain_lib.link_note_in_daily(existing, timestamp=now_ts)
        if verbose:
            _step(f"{prefix}♻️  al verwerkt → [[{existing}]]")
        if _should_archive_after_save(
            mail, no_archive=no_archive, single_mode=single_mode
        ):
            try:
                archive_mail(
                    message_id,
                    mail.get("account", ""),
                    mailbox=mail.get("mailbox", ""),
                )
                _timed("archive")
                if verbose:
                    _step(f"{prefix}📦 Archive... OK")
            except Exception:
                pass
        result = {
            "status": "duplicate",
            "slug": existing,
            "subject": subject,
            "direction": infer_mail_direction(mail),
            "path": str(find_note_path_by_slug(existing) or ""),
            "mailbox_type": mail["mailbox_type"],
        }
        _write_log(result)
        return result

    if verbose:
        _step(f"{prefix}🔍 Dedup... nieuw")
    mail["body"] = fetch_mail_body(
        message_id,
        account=mail.get("account", ""),
        mailbox=mail.get("mailbox", ""),
    )
    _timed("body")

    create_task = should_create_follow_up_task(
        mail,
        force_task=force_task,
        mailbox_type=mail["mailbox_type"],
    )
    result = create_mail_note(
        mail,
        create_task=create_task,
        extra_thread_slugs=extra_thread_slugs,
    )
    _timed("note")
    if verbose:
        size_suffix = (
            f" ({result['note_size']})" if result.get("note_size") else ""
        )
        _step(f"{prefix}📝 → [[{result['slug']}]]{size_suffix}")
        if result.get("task_slug"):
            _step(f"{prefix}📋 Task → [[{result['task_slug']}]]")

    if _should_archive_after_save(mail, no_archive=no_archive, single_mode=single_mode):
        archive_result = archive_mail(
            message_id,
            mail.get("account", ""),
            mailbox=mail.get("mailbox", ""),
        )
        _timed("archive")
        result["archived"] = archive_result
        result["archive_account"] = mail.get("account", "")
        result["archive_mailbox"] = mail.get("mailbox", "")
        if archive_result != "OK":
            result["archive_error"] = archive_result
            if verbose:
                _step(f"{prefix}⚠️  Archive overgeslagen: {archive_result}")
        else:
            if verbose:
                _step(f"{prefix}📦 Archive... OK")

    result["status"] = "saved"
    result["mailbox_type"] = mail["mailbox_type"]
    _write_log(result)
    return result


def calendar_invite_selected(mail: dict[str, Any]) -> bool:
    return _mailbox_type(mail) == "deleted" and is_calendar_invite(mail)


def process_selected_mail_records(
    mails: list[dict[str, Any]],
    *,
    force_task: bool = False,
    no_archive: bool = False,
    force_conversation: bool = False,
    verbose: bool = True,
) -> dict[str, Any]:
    """Use selected mails as conversation roots without scanning whole accounts."""
    results: list[dict[str, Any]] = []
    thread_updated = 0
    total = len(mails)
    for index, mail in enumerate(mails, 1):
        mailbox_type = _mailbox_type(mail)
        subject = mail.get("subject", "?")[:60]
        sender = mail.get("sender_email", "?")
        if verbose:
            if calendar_invite_selected(mail):
                marker = "📅"
            elif mailbox_type == "sent":
                marker = "📤"
            else:
                marker = "📧"
            _step(f'[{index}/{total}] {marker} "{subject}" van {sender}')
        outcome = save_mail_with_conversation(
            mail,
            force_task=force_task,
            no_archive=no_archive,
            force_conversation=force_conversation,
            allow_subject_fallback=force_conversation,
            verbose=verbose,
        )
        results.extend(outcome["results"])
        thread_updated += int(outcome.get("thread_updated") or 0)
    return {
        "mode": "selection",
        "reason": "selected-conversation",
        "results": results,
        "thread_updated": thread_updated,
    }


def save_mail_with_conversation(
    selected_mail: dict[str, Any],
    *,
    force_task: bool = False,
    no_archive: bool = False,
    force_conversation: bool = False,
    allow_subject_fallback: bool = True,
    verbose: bool = True,
) -> dict[str, Any]:
    """Save one selected mail, auto-expanding safe conversations."""
    hints = conversation_fetch_hints(
        selected_mail, force_subject=force_conversation
    )
    lookup_ids = conversation_lookup_ids(
        selected_mail, force=force_conversation
    )
    existing_thread_by_id = existing_thread_slug_map(lookup_ids)
    existing_thread_slugs = list(dict.fromkeys(existing_thread_by_id.values()))
    missing_lookup_ids = [
        message_id
        for message_id in lookup_ids
        if clean_message_id(message_id).lower() not in existing_thread_by_id
    ]
    should_lookup_conversation = bool(
        force_conversation
        or missing_lookup_ids
        or (hints["include_subject"] and not existing_thread_slugs)
    )
    if verbose:
        _step("🧵 Checking conversation...")

    selection = {
        "is_conversation": False,
        "reason": "single",
        "mails": [selected_mail],
        "reasons": [],
    }
    if should_lookup_conversation:
        fetch_lookup_ids = lookup_ids if force_conversation else missing_lookup_ids
        if fetch_lookup_ids:
            candidates = fetch_conversation_candidates(
                selected_mail,
                fetch_lookup_ids,
                hints["subject_hint"],
                include_subject=False,
            )
            selection = select_conversation_mails(
                selected_mail,
                candidates,
                force=force_conversation,
            )

        if (
            not selection["is_conversation"]
            and allow_subject_fallback
            and hints["include_subject"]
            and (force_conversation or not existing_thread_slugs)
        ):
            candidates = fetch_conversation_candidates(
                selected_mail,
                fetch_lookup_ids,
                hints["subject_hint"],
                include_subject=True,
                search_other_accounts=force_conversation,
            )
            selection = select_conversation_mails(
                selected_mail,
                candidates,
                force=force_conversation,
            )
    _timed("conversation")

    if not selection["is_conversation"]:
        if verbose:
            _step("   Geen betrouwbare conversation gevonden")
            _step("")
        result = process_mail_record(
            selected_mail,
            force_task=force_task,
            no_archive=no_archive,
            single_mode=True,
            extra_thread_slugs=existing_thread_slugs,
            verbose=verbose,
        )
        updated = 0
        if existing_thread_slugs and result.get("slug"):
            timeline_entries = [
                _thread_entry_from_slug(slug) for slug in existing_thread_slugs
            ]
            timeline_entries.append(
                {
                    "slug": result["slug"],
                    "direction": result.get("direction")
                    or infer_mail_direction(selected_mail),
                    "path": result.get("path", ""),
                }
            )
            updated = update_conversation_thread_timeline(timeline_entries)
            if verbose and updated:
                _step(f"   🔗 Thread timeline bijgewerkt in {updated} notes")
        return {
            "mode": "single",
            "reason": "existing-thread" if existing_thread_slugs else selection["reason"],
            "results": [result],
            "thread_updated": updated,
        }

    mails = selection["mails"]
    if verbose:
        _step(f"   Conversation: {len(mails)} mails ({selection['reason']})")
        _step("")

    selected_id = clean_message_id(selected_mail.get("message_id", "")).lower()
    thread_slugs: list[str] = list(existing_thread_slugs)
    results: list[dict[str, Any]] = []
    timeline_entries: list[dict[str, Any]] = [
        _thread_entry_from_slug(slug) for slug in existing_thread_slugs
    ]
    for index, mail in enumerate(mails, 1):
        subject = mail.get("subject", "?")[:60]
        sender = mail.get("sender_email", "?")
        if verbose:
            _step(f"[{index}/{len(mails)}] {subject} van {sender}")
        mail_id = clean_message_id(mail.get("message_id", "")).lower()
        result = process_mail_record(
            mail,
            force_task=force_task and mail_id == selected_id,
            no_archive=no_archive,
            single_mode=False,
            extra_thread_slugs=list(thread_slugs),
            verbose=verbose,
            prefix="      ",
        )
        slug = result.get("slug")
        if slug and slug not in thread_slugs:
            thread_slugs.append(slug)
        if slug:
            timeline_entries.append(
                {
                    "slug": slug,
                    "direction": result.get("direction") or infer_mail_direction(mail),
                    "path": result.get("path", ""),
                }
            )
        results.append(result)

    updated = update_conversation_thread_timeline(timeline_entries)
    if verbose and updated:
        _step(f"   🔗 Thread timeline bijgewerkt in {updated} notes")

    return {
        "mode": "conversation",
        "reason": selection["reason"],
        "results": results,
        "thread_updated": updated,
    }


def save_selected_mail(
    *,
    force_task: bool = False,
    no_archive: bool = False,
    single: bool = False,
    force_conversation: bool = False,
    client: str = MAIL_CLIENT_APPLE,
    selected_mails: list[dict[str, Any]] | None = None,
    verbose: bool = True,
) -> dict[str, Any]:
    """Save selected mail messages, auto-expanding safe Apple Mail conversations."""
    if verbose:
        _step("📨 Fetching selected mail...")
    resolved_client = client
    if selected_mails is None:
        resolved_client, selected_mails = get_selected_headers_for_client(client)
    for mail in selected_mails:
        mail.setdefault("mail_client", resolved_client)
    _timed("selected")

    if verbose:
        if len(selected_mails) > 1:
            _step(f"   {len(selected_mails)} mails geselecteerd")
        else:
            selected_mail = selected_mails[0]
            _step(f"   {selected_mail.get('subject', '?')}")
            _step(f"   From: {selected_mail.get('sender_email', '?')}")
        _step("")

    selected_mail = selected_mails[0]
    if single:
        if verbose:
            _step("📝 Single mail mode")
        with mail_client_context(resolved_client):
            result = process_mail_record(
                selected_mail,
                force_task=force_task,
                no_archive=no_archive,
                single_mode=True,
                verbose=verbose,
            )
        return {"mode": "single", "reason": "forced", "results": [result]}

    if len(selected_mails) > 1:
        with mail_client_context(resolved_client):
            return process_selected_mail_records(
                selected_mails,
                force_task=force_task,
                no_archive=no_archive,
                force_conversation=force_conversation,
                verbose=verbose,
            )

    with mail_client_context(resolved_client):
        return save_mail_with_conversation(
            selected_mail,
            force_task=force_task,
            no_archive=no_archive,
            force_conversation=force_conversation,
            verbose=verbose,
        )


# ── Batch Processing ───────────────────────────────────────────────────────


def _prune_dedup_cache(max_age_days: int = 90) -> int:
    """Remove dedup cache entries older than max_age_days. Returns pruned count."""
    try:
        if not _DEDUP_CACHE.exists():
            return 0
        cache = json.loads(_DEDUP_CACHE.read_text())
        if not cache:
            return 0
        from datetime import datetime, timedelta

        cutoff = datetime.now() - timedelta(days=max_age_days)
        pruned = 0
        new_cache = {}
        for msg_id, slug in cache.items():
            # Slug format: YYYYMMDD-HHMM-... — extract date from first 8 chars
            try:
                entry_date = datetime.strptime(slug[:8], "%Y%m%d")
                if entry_date >= cutoff:
                    new_cache[msg_id] = slug
                else:
                    pruned += 1
            except (ValueError, IndexError):
                new_cache[msg_id] = slug  # keep entries with unparseable dates
        if pruned > 0:
            _DEDUP_CACHE.write_text(json.dumps(new_cache))
        return pruned
    except Exception:
        return 0


def batch_process(as_json: bool = False, since_days: int = 7) -> list[dict]:
    """Process inbox/sent/deleted mails sequentially. Stops on error.

    since_days: only process mails newer than this many days (default 7).
                Use 0 for no limit (process all).
    """
    import time

    # Prune stale dedup cache entries on each batch run
    pruned = _prune_dedup_cache()
    if pruned > 0 and not as_json:
        _step(f"Dedup cache: {pruned} entries ouder dan 90 dagen verwijderd")

    if not is_mail_running():
        if not as_json:
            _step("⚠️  Mail.app is niet actief — overgeslagen")
        return []

    if since_days > 0 and not as_json:
        from datetime import timedelta

        cutoff_date = (datetime.now() - timedelta(days=since_days)).strftime("%Y-%m-%d")
        _step(f"📅 Cutoff: mails vanaf {cutoff_date}")

    # Fetch all mailboxes in one pass (7 AppleScript calls instead of 21)
    if not as_json:
        _step("📨 Mails ophalen...")
    all_mailboxes = fetch_all_mailboxes(since_days=since_days)

    all_results: list[dict] = []
    mailbox_order = [
        ("inbox", "📧", "INBOX"),
        ("sent", "📤", "SENT"),
        ("deleted", "📅", "DELETED (calendar)"),
    ]

    for mailbox_type, icon, label in mailbox_order:
        mails = all_mailboxes.get(mailbox_type, [])

        # Filter deleted to calendar invites only
        if mailbox_type == "deleted":
            mails = [m for m in mails if is_calendar_invite(m)]

        # Pre-filter sent/deleted duplicates (they can't be archived, so skip silently)
        skipped_dupes = 0
        if mailbox_type in ("sent", "deleted") and mails:
            new_mails = []
            for m in mails:
                cached_slug = is_duplicate(m.get("message_id", ""), mail=m)
                if cached_slug:
                    skipped_dupes += 1
                    all_results.append(
                        {
                            "status": "duplicate",
                            "slug": cached_slug,
                            "mailbox_type": mailbox_type,
                            "batch": True,
                        }
                    )
                else:
                    new_mails.append(m)
            mails = new_mails

        if not mails:
            if skipped_dupes and not as_json:
                _step(f"\n── {label} ── {skipped_dupes} al verwerkt, 0 nieuw")
            continue

        should_archive = mailbox_type == "inbox"

        suffix = f" (+{skipped_dupes} al verwerkt)" if skipped_dupes else ""
        if not as_json:
            _step(f"\n── {label} ({len(mails)} mails{suffix}) ──")

        for i, mail in enumerate(mails, 1):
            t0 = time.time()
            subject = mail.get("subject", "?")[:60]
            sender = mail.get("sender_email", "?")
            is_flagged = mail.get("is_flagged", False)
            flag_marker = " ⚑" if is_flagged else ""

            if not as_json:
                _step(
                    f'[{i}/{len(mails)}] {icon} "{subject}" van {sender}{flag_marker}'
                )

            # 1. Dedup
            existing = is_duplicate(mail["message_id"], mail=mail)
            if existing:
                if not as_json:
                    _step(f"      ♻️  al verwerkt → [[{existing}]]")
                # Archive dupes from inbox so they leave the inbox
                if should_archive:
                    try:
                        archive_mail(
                            mail["message_id"],
                            mail["account"],
                            mailbox=mail.get("mailbox", ""),
                        )
                        if not as_json:
                            _step("      📦 Archive... OK")
                    except Exception:
                        pass
                all_results.append(
                    {
                        "status": "duplicate",
                        "slug": existing,
                        "mailbox_type": mailbox_type,
                        "batch": True,
                    }
                )
                continue

            # 2. Fetch body
            if not as_json:
                _step("      🔍 Dedup... nieuw")
            mail["body"] = fetch_mail_body(
                mail["message_id"],
                account=mail.get("account", ""),
                mailbox=mail.get("mailbox", ""),
            )

            # 3. Create note (+ task if flagged inbox)
            create_task = should_create_follow_up_task(
                mail, mailbox_type=mailbox_type
            )
            try:
                result = create_mail_note(mail, create_task=create_task)
            except Exception as e:
                _step(f"      ❌ Note creation failed: {e}")
                _write_log(
                    {
                        "status": "error",
                        "message": str(e),
                        "subject": subject,
                        "mailbox_type": mailbox_type,
                        "batch": True,
                    }
                )
                raise RuntimeError(
                    f'Batch gestopt bij [{i}/{len(mails)}] "{subject}": {e}'
                )

            if not as_json:
                size_suffix = (
                    f" ({result['note_size']})" if result.get("note_size") else ""
                )
                _step(f"      📝 → [[{result['slug']}]]{size_suffix}")

            if result.get("task_slug") and not as_json:
                _step(f"      📋 Task → [[{result['task_slug']}]]")

            # 4. Archive (inbox only)
            if should_archive:
                try:
                    archive_result = archive_mail(
                        mail["message_id"],
                        mail["account"],
                        mailbox=mail.get("mailbox", ""),
                    )
                except Exception as e:
                    archive_result = f"ERROR:{e}"

                result["archived"] = archive_result
                result["archive_account"] = mail.get("account", "")
                result["archive_mailbox"] = mail.get("mailbox", "")
                if archive_result != "OK":
                    result["archive_error"] = archive_result

                if not as_json:
                    if archive_result == "OK":
                        _step("      📦 Archive... OK")
                    else:
                        _step(f"      ⚠️  Archive overgeslagen: {archive_result}")

            elapsed = time.time() - t0
            if not as_json:
                size_prefix = (
                    f"{result['note_size']}, " if result.get("note_size") else ""
                )
                _step(f"      ✅ ({size_prefix}{elapsed:.1f}s)")

            # 5. Log
            result["status"] = "saved"
            result["mailbox_type"] = mailbox_type
            result["batch"] = True
            _write_log(result)
            all_results.append(result)

    return all_results


# ── Main ─────────────────────────────────────────────────────────────────────


_LOG_PATH = brain_lib.ROOT / "context" / "observability" / "save-mail.jsonl"
_t0 = 0.0
_timings: dict[str, float] = {}


def _step(msg: str) -> None:
    """Print a progress step (flushed for real-time Raycast output)."""
    print(msg, flush=True)


def _file_size(path: Path) -> int | None:
    try:
        return path.stat().st_size
    except OSError:
        return None


def format_data_size(num_bytes: int) -> str:
    """Format a byte count for Raycast output."""
    if num_bytes < 1024:
        return f"{num_bytes} B"
    if num_bytes < 1024 * 1024:
        return f"{num_bytes / 1024:.1f} KB"
    return f"{num_bytes / (1024 * 1024):.1f} MB"


def _total_note_size(results: list[dict[str, Any]]) -> str:
    total = sum(
        int(result["note_bytes"])
        for result in results
        if isinstance(result.get("note_bytes"), int)
    )
    return format_data_size(total) if total else ""


def _timed(label: str) -> float:
    """Record timing for a step. Returns elapsed ms since last call."""
    import time

    if _t0 <= 0:
        return 0.0
    now = time.time()
    elapsed = (now - _t0) * 1000
    _timings[label] = elapsed
    return elapsed


def _write_log(entry: dict) -> None:
    """Append a JSON log entry to the save-mail log."""
    try:
        _LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        log = {**entry, "timings": dict(_timings), "ts": datetime.now().isoformat()}
        with open(_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(json.dumps(log, default=str) + "\n")
    except Exception:
        pass


def main() -> None:
    import argparse
    import time

    global _t0
    _t0 = time.time()

    parser = argparse.ArgumentParser(description="Save selected mail to vault")
    parser.add_argument("--task", action="store_true", help="Create follow-up task")
    parser.add_argument("--no-archive", action="store_true", help="Skip archiving")
    parser.add_argument("--json", action="store_true", help="JSON output")
    parser.add_argument(
        "--client",
        choices=MAIL_CLIENTS,
        default=MAIL_CLIENT_AUTO,
        help="Mail client to read from (default: auto)",
    )
    parser.add_argument(
        "--single",
        action="store_true",
        help="Force old single-message behavior",
    )
    parser.add_argument(
        "--conversation",
        action="store_true",
        help="Force conversation lookup even when confidence is low",
    )
    args = parser.parse_args()
    verbose = not args.json

    if args.single and args.conversation:
        parser.error("--single and --conversation cannot be used together")

    if verbose:
        _step(f"▶ Script: {Path(__file__).resolve()}")
        _step(f"🧾 Log: {_LOG_PATH.resolve()}")

    try:
        outcome = save_selected_mail(
            force_task=args.task,
            no_archive=args.no_archive,
            single=args.single,
            force_conversation=args.conversation,
            client=args.client,
            verbose=verbose,
        )
        total_ms = _timed("total")
        if args.json:
            print(json.dumps(outcome, default=str))
        else:
            results = outcome["results"]
            saved = [r for r in results if r.get("status") == "saved"]
            dupes = [r for r in results if r.get("status") == "duplicate"]
            _step("━" * 40)
            if outcome["mode"] == "conversation":
                size = _total_note_size(saved)
                size_suffix = f", {size}" if size else ""
                _step(
                    f"✅ Conversation klaar: {len(saved)} opgeslagen, "
                    f"{len(dupes)} al verwerkt{size_suffix} "
                    f"({total_ms / 1000:.1f}s)"
                )
            elif outcome["mode"] == "selection":
                size = _total_note_size(saved)
                size_suffix = f", {size}" if size else ""
                _step(
                    f"✅ Selectie klaar: {len(saved)} opgeslagen, "
                    f"{len(dupes)} al verwerkt{size_suffix} "
                    f"({total_ms / 1000:.1f}s)"
                )
            else:
                result = results[0]
                if result.get("status") == "duplicate":
                    _step(f"✅ Al verwerkt: [[{result['slug']}]]")
                else:
                    size_prefix = (
                        f"{result['note_size']}, "
                        if result.get("note_size")
                        else ""
                    )
                    _step(
                        f"✅ [[{result['slug']}]]  "
                        f"({size_prefix}{total_ms / 1000:.1f}s)"
                    )

    except RuntimeError as e:
        error_msg = str(e)
        _timed("error")
        _write_log({"status": "error", "message": error_msg})
        if args.json:
            print(json.dumps({"status": "error", "message": error_msg}))
        else:
            _step(f"\n❌ {error_msg}")
            if (
                "No selected mail found in Outlook or Apple Mail" in error_msg
                or "No selected Outlook mail." in error_msg
                or "Copied clipboard content is not an Outlook message." in error_msg
            ):
                return
        sys.exit(1)
    except Exception as e:
        _timed("error")
        _write_log({"status": "error", "message": str(e)})
        if args.json:
            print(json.dumps({"status": "error", "message": str(e)}))
        else:
            _step(f"\n❌ {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
