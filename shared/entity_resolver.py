#!/usr/bin/env python3
"""Shared entity resolution: email→entity slug, topic detection.

Used by save_mail.py, save_message.py, create_mail_notes.py.
Provides O(1) lookups via one-shot SQLite cache load.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path

import brain_lib

VAULT_ROOT = brain_lib.cfg.vault_root

# ── Self Emails ─────────────────────────────────────────────────────────────

SELF_EMAILS = {
    "viggomulders@outlook.com",
    "vruqq@outlook.com",
    "viggo.meesters@mccoy-partners.com",
    "viggo.meesters@groningen.nl",
    "ricardoviggomulders@gmail.com",
    "viggomeesters@icloud.com",
    "viggomeesters@protonmail.com",
    "viqqo@pm.me",
    "1990_kees@live.nl",
    "ricardo001@live.nl",
}

# ── Known Email → Entity Slug Mappings ──────────────────────────────────────

ENTITY_MAP: dict[str, str] = {
    "aebuyer services@aliexpress.com": "aliexpress",
    "aebuyerservices@aliexpress.com": "aliexpress",
    "cf.topdesk@uu.nl": "uu-it-support",
    "community@email-ws.withings.com": "withings",
    "discover@airbnb.com": "airbnb",
    "do-not-reply@uipath.com": "uipath",
    "donotreply@sappartnerupdate.com": "sappartnerupdate",
    "google-maps-noreply@google.com": "google",
    "kevin@getdex.com": "getdex",
    "message@info.aliexpress.com": "aliexpress",
    "no-reply@boldsmartlock.com": "boldsmartlock",
    "no-reply@rs.email.nextdoor.nl": "nextdoor",
    "no-reply@chargemap.com": "chargemap",
    "no-reply@mail.proton.me": "proton",
    "no-reply@todoist.com": "todoist",
    "no-reply@update.bunq.com": "bunq",
    "noreply-photos@google.com": "google",
    "noreply@email.openai.com": "openai",
    "noreply@glassdoor.com": "glassdoor",
    "noreply@google.com": "google",
    "notificatie@edm.postnl.nl": "postnl",
}

RELAY_MAP: dict[str, str] = {
    "instant-gaming": "instant-gaming",
}

# Generic email providers — NEVER use for domain-based entity matching
GENERIC_PROVIDERS = frozenset(
    {
        "gmail.com",
        "googlemail.com",
        "outlook.com",
        "outlook.nl",
        "hotmail.com",
        "hotmail.nl",
        "live.com",
        "live.nl",
        "icloud.com",
        "me.com",
        "mac.com",
        "protonmail.com",
        "proton.me",
        "pm.me",
        "yahoo.com",
        "yahoo.nl",
        "ymail.com",
        "msn.com",
        "msn.nl",
        "kpnmail.nl",
        "ziggo.nl",
        "xs4all.nl",
        "upcmail.nl",
        "chello.nl",
        "planet.nl",
        "home.nl",
        "casema.nl",
        "hetnet.nl",
        "quicknet.nl",
        "solcon.nl",
        "tele2.nl",
        "aol.com",
        "mail.com",
        "gmx.com",
        "gmx.net",
    }
)

# Prefixes to strip from email local parts (info@, noreply@, etc.)
GENERIC_PREFIXES = frozenset(
    {
        "info",
        "contact",
        "noreply",
        "no-reply",
        "no.reply",
        "hello",
        "support",
        "admin",
        "sales",
        "billing",
        "help",
        "service",
        "webmaster",
        "postmaster",
        "mail",
        "team",
        "office",
        "enquiries",
        "general",
    }
)

MAIL_DOMAIN_PREFIXES = frozenset(
    {
        "email",
        "mail",
        "mailer",
        "notify",
        "notifications",
        "newsletter",
        "news",
        "reply",
        "replies",
    }
)

COMMON_TLDS = frozenset(
    {
        "com",
        "nl",
        "org",
        "net",
        "io",
        "co",
        "de",
        "be",
        "eu",
        "uk",
        "app",
    }
)

# ── Cached Entity Email Lookups (built once from SQLite) ────────────────────

_email_cache: dict[str, str] = {}
_domain_cache: dict[str, str] = {}
_known_slugs: set[str] = set()
_entity_topics_cache: dict[str, list[str]] = {}
_all_topics: set[str] = set()
_topic_pattern: re.Pattern[str] | None = None
_cache_loaded = False


@dataclass(frozen=True)
class EntityResolution:
    slug: str
    direction: str
    source: str
    confidence: float


def _load_entity_cache() -> None:
    """One-shot load: email→entity, domain→entity, entity→topics, all topics."""
    global _cache_loaded, _topic_pattern
    if _cache_loaded:
        return
    _cache_loaded = True
    try:
        import sqlite3

        db_path = VAULT_ROOT / ".brain-vault-index.sqlite"
        if not db_path.exists():
            return
        conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
        try:
            conn.execute("PRAGMA query_only = ON")

            rows = conn.execute(
                "SELECT slug, category, json_extract(frontmatter_json, '$.email'), "
                "json_extract(frontmatter_json, '$.emails'), topics "
                "FROM notes WHERE type = 'entity' "
                "ORDER BY CASE WHEN category = 'company' THEN 0 ELSE 1 END"
            ).fetchall()
            for slug, category, email, emails_json, topics in rows:
                _known_slugs.add(slug)
                if email:
                    email_lower = email.lower().strip()
                    _email_cache[email_lower] = slug
                    domain = email_lower.split("@")[-1]
                    if domain not in GENERIC_PROVIDERS:
                        parts = domain.split(".")
                        d_key = parts[-2] if len(parts) >= 3 else parts[0]
                        if d_key:
                            if d_key not in _domain_cache or slug == d_key:
                                _domain_cache[d_key] = slug
                if emails_json:
                    try:
                        addrs = (
                            json.loads(emails_json)
                            if isinstance(emails_json, str)
                            else emails_json
                        )
                        if isinstance(addrs, list):
                            for entry in addrs:
                                addr = (
                                    entry.get("value", "")
                                    if isinstance(entry, dict)
                                    else str(entry)
                                )
                                if addr:
                                    _email_cache[addr.lower().strip()] = slug
                    except (json.JSONDecodeError, TypeError):
                        pass
                if topics and topics != "[]":
                    try:
                        parsed = (
                            json.loads(topics) if isinstance(topics, str) else topics
                        )
                        if isinstance(parsed, list) and parsed:
                            _entity_topics_cache[slug] = parsed
                    except (json.JSONDecodeError, TypeError):
                        pass

            raw_topics: set[str] = set()
            topic_rows = conn.execute(
                "SELECT DISTINCT topics FROM notes "
                "WHERE topics <> '[]' AND topics <> '' AND topics IS NOT NULL"
            ).fetchall()
            for (t,) in topic_rows:
                try:
                    parsed = json.loads(t) if isinstance(t, str) else t
                    if isinstance(parsed, list):
                        raw_topics.update(parsed)
                except (json.JSONDecodeError, TypeError):
                    pass

            for topic in raw_topics:
                if (
                    len(topic) >= 4
                    and not topic.isdigit()
                    and not re.match(r"^\d{2}:\d{2}", topic)
                ):
                    _all_topics.add(topic)
            if _all_topics:
                escaped = sorted(
                    (re.escape(t) for t in _all_topics), key=len, reverse=True
                )
                _topic_pattern = re.compile(
                    r"\b(" + "|".join(escaped) + r")\b", re.IGNORECASE
                )
        finally:
            conn.close()
    except Exception:
        pass


def resolve_entity(
    email: str, to_emails: str | None = None, *, fallback_fn=None
) -> tuple[str, str]:
    """Resolve email to (entity_slug, direction). Fast path first."""
    result = resolve_entity_details(email, to_emails, fallback_fn=fallback_fn)
    return result.slug, result.direction


def resolve_entity_details(
    email: str, to_emails: str | None = None, *, fallback_fn=None
) -> EntityResolution:
    """Resolve email to entity metadata, including source and confidence.

    If *fallback_fn* is provided and no cached/mapped match is found, it is
    called with the email address.  It should return a ``(slug, direction)``
    tuple or ``None`` to fall through to the default domain-based fallback.
    """
    _load_entity_cache()
    email_lower = email.lower().strip()

    if "@" not in email_lower:
        return EntityResolution("unknown-sender", "received", "missing-email", 0.0)

    if email_lower in SELF_EMAILS:
        if to_emails:
            first_to = to_emails.split(",")[0].strip().lower()
            if first_to and first_to not in SELF_EMAILS:
                recipient = resolve_entity_details(first_to, fallback_fn=fallback_fn)
                return EntityResolution(
                    recipient.slug,
                    "sent",
                    f"sent-recipient:{recipient.source}",
                    min(recipient.confidence, 0.95),
                )
        return EntityResolution("viggo-meesters", "sent", "self-email", 1.0)

    if email_lower in ENTITY_MAP:
        return EntityResolution(ENTITY_MAP[email_lower], "received", "known-email-map", 1.0)

    if email_lower in _email_cache:
        return EntityResolution(_email_cache[email_lower], "received", "vault-email", 1.0)

    if "privaterelay.appleid.com" in email_lower:
        for key, slug in RELAY_MAP.items():
            if key in email_lower:
                return EntityResolution(slug, "received", "apple-relay-map", 0.9)
        slug = _derive_entity_from_relay_local(email_lower.split("@", 1)[0])
        if slug:
            return EntityResolution(slug, "received", "apple-relay-derived", 0.75)

    domain = email_lower.split("@")[-1]
    local = email_lower.split("@")[0]
    if domain == "icloud.com" and "_at_" in local and email_lower not in SELF_EMAILS:
        slug = _derive_entity_from_relay_local(local)
        if slug:
            if slug in _domain_cache:
                return EntityResolution(_domain_cache[slug], "received", "vault-domain-relay", 0.9)
            return EntityResolution(slug, "received", "icloud-relay-derived", 0.75)

    if domain not in GENERIC_PROVIDERS and _is_likely_person(local):
        slug = _derive_entity_from_local_part(local)
        if email_lower not in SELF_EMAILS:
            _ensure_entity_exists(slug, email_lower, domain, category="person")
        return EntityResolution(slug, "received", "derived-person-local-part", 0.7)

    # Domain-based cache lookup (only for non-generic providers)
    if domain not in GENERIC_PROVIDERS:
        parts = domain.split(".")
        candidates = []
        if len(parts) >= 3:
            candidates.append(parts[-2])
        candidates.append(parts[0])
        for d_key in candidates:
            if d_key in _domain_cache:
                return EntityResolution(_domain_cache[d_key], "received", "vault-domain", 0.9)

    # Try custom fallback before smart parsing
    if fallback_fn is not None:
        result = fallback_fn(email_lower)
        if result is not None:
            slug, direction = result
            return EntityResolution(slug, direction, "custom-fallback", 0.6)

    # Smart fallback: derive entity slug from email address
    slug = _derive_entity_from_email(email_lower)

    # Auto-create entity note if it doesn't exist in vault (skip self emails)
    if email_lower not in SELF_EMAILS:
        _ensure_entity_exists(slug, email_lower, domain)

    return EntityResolution(slug, "received", "derived-email", 0.55)


def _derive_entity_from_email(email: str) -> str:
    """Derive a reasonable entity slug from an email address.

    Custom domain: info@depaarsekeizerin.nl → depaarsekeizerin
    Generic provider with name: viggo.meesters@gmail.com → viggo-meesters
    Generic provider without name: 1990_kees@live.nl → kees
    Plus addressing: viggo+test@gmail.com → viggo
    """
    local, domain = email.split("@", 1)

    # Strip plus addressing (viggo+test → viggo)
    local = local.split("+")[0]

    if domain in GENERIC_PROVIDERS:
        # Parse the local part as a person name
        # Strip leading/trailing digits (e.g. "1990_kees" → "kees", "romy123" → "romy")
        cleaned = re.sub(r"^\d+[._-]?", "", local)
        cleaned = re.sub(r"[._-]?\d+$", "", cleaned)
        if not cleaned:
            cleaned = local
        # If local part is a generic prefix on a generic provider, not useful
        if cleaned.lower() in GENERIC_PREFIXES:
            return domain.split(".")[0].lower()
        # Split on . _ - and join as slug
        parts = re.split(r"[._-]+", cleaned)
        parts = [p for p in parts if p and not p.isdigit()]
        if parts:
            return "-".join(parts).lower()
        return cleaned.lower()

    # Custom domain: use domain as entity slug
    domain_parts = domain.split(".")
    if len(domain_parts) >= 3:
        slug = domain_parts[-2]
    else:
        slug = domain_parts[0]

    return slug.lower().replace("_", "-")


def _derive_entity_from_local_part(local: str) -> str:
    """Derive a person slug from an email local part."""
    local = local.split("+")[0]
    cleaned = re.sub(r"^\d+[._-]?", "", local)
    cleaned = re.sub(r"[._-]?\d+$", "", cleaned)
    parts = re.split(r"[._-]+", cleaned)
    parts = [p for p in parts if p and not p.isdigit()]
    return "-".join(parts).lower()


def _derive_entity_from_relay_local(local: str) -> str:
    """Derive an entity slug from Apple relay local parts.

    Examples:
    noreply_at_email_openai_com_x@icloud.com → openai
    noreply_at_email_hansanders_nl_x@icloud.com → hansanders
    """
    if "_at_" not in local:
        return ""

    after_at = local.split("_at_", 1)[1]
    labels = [p for p in re.split(r"[_-]+", after_at) if p]
    if not labels:
        return ""

    tld_index = next(
        (i for i, p in enumerate(labels) if i > 0 and p in COMMON_TLDS), None
    )
    domain_labels = labels[:tld_index] if tld_index is not None else labels
    domain_labels = [p for p in domain_labels if p not in MAIL_DOMAIN_PREFIXES]

    if domain_labels:
        return domain_labels[-1].lower()
    return labels[0].lower()


def _humanize_company_slug(slug: str) -> str:
    """Try to make a company slug human-readable.

    depaarsekeizerin → Depaarsekeizerin (no word boundary detection)
    mccoy-partners → Mccoy Partners
    coolblue → Coolblue
    """
    return slug.replace("-", " ").title()


def _is_likely_person(local: str) -> bool:
    """Heuristic: does the local part look like a person name (e.g. jan.de.vries)?"""
    parts = re.split(r"[._-]+", local)
    parts = [p for p in parts if p and not p.isdigit() and len(p) > 1]
    return len(parts) >= 2 and parts[0].lower() not in GENERIC_PREFIXES


def _ensure_entity_exists(
    slug: str, email: str, domain: str, *, category: str | None = None
) -> None:
    """Auto-create a minimal entity note if no entity with this slug exists."""
    entity_dir = brain_lib.cfg.vault_entities
    entity_path = entity_dir / f"{slug}.md"
    if entity_path.exists():
        return

    # Check if slug exists as a known entity
    if slug in _known_slugs:
        return

    try:
        import vault_note_writer as vnw
    except ImportError:
        return

    from datetime import datetime

    now = datetime.now()
    local = email.split("@")[0].split("+")[0]

    # Determine category: generic provider → person, custom domain → company
    if category == "person" or domain in GENERIC_PROVIDERS:
        category = "person"
        cleaned = re.sub(r"^\d+[._-]?", "", local)
        cleaned = re.sub(r"[._-]?\d+$", "", cleaned)
        name_parts = re.split(r"[._-]+", cleaned)
        name_parts = [p.capitalize() for p in name_parts if p and not p.isdigit()]
        title = " ".join(name_parts) if name_parts else slug.replace("-", " ").title()
    else:
        category = "company"
        title = _humanize_company_slug(slug)

    fm = {
        "type": "entity",
        "category": category,
        "created": now.strftime("%Y-%m-%d"),
        "slug": slug,
        "timestamp": now.strftime("%Y%m%d-%H%M"),
        "area": "self",
        "title": vnw.escape_title(title),
        "email": email,
        "source": "auto-created-from-mail",
    }

    body = f"# {title}\n"

    entity_dir.mkdir(parents=True, exist_ok=True)
    vnw.write_vault_note(entity_path, fm, body)
    brain_lib.link_note_in_daily(slug, now.strftime("%Y%m%d-%H%M"))

    # Add to caches so subsequent lookups hit
    _known_slugs.add(slug)
    _email_cache[email.lower()] = slug
    if domain not in GENERIC_PROVIDERS:
        parts = domain.split(".")
        d_key = parts[-2] if len(parts) >= 3 else parts[0]
        if d_key:
            _domain_cache.setdefault(d_key, slug)


def suggest_topics(entity_slug: str, subject: str) -> list[str]:
    """Auto-detect topics from entity profile and subject keywords."""
    _load_entity_cache()
    topics: set[str] = set()

    if entity_slug in _entity_topics_cache:
        topics.update(_entity_topics_cache[entity_slug])

    if _topic_pattern and subject:
        topics.update(_topic_pattern.findall(subject.lower()))

    return sorted(topics)
