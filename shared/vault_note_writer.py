"""Shared vault note writer — frontmatter rendering, slug generation, validation.

Infrastructure module used by save_mail.py, save_message.py, save_bookmark.py,
and quick_task.py. Handles HOW notes are written; callers decide WHAT to write.

Zero side effects on import. No global state.
"""

from __future__ import annotations

import logging
import os
import re
import tempfile
from datetime import date, datetime
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlencode, urlparse, urlunparse

try:
    from contract_paths import contract_path
except ImportError:  # pragma: no cover - package import fallback
    from .contract_paths import contract_path

logger = logging.getLogger(__name__)

# ── Schema constants ────────────────────────────────────────────────────────
# Hardcoded defaults; overridden by life-os-schema.yaml if available.

VALID_TYPES = frozenset(
    {
        "interaction",
        "reference",
        "task",
        "health",
        "chore",
        "entity",
        "project",
        "anniversary",
        "channel",
        "routine",
        "purchase",
        "entry",
        "context",
        "architecture",
    }
)

VALID_AREAS = frozenset({"work", "home", "self", "social"})

STATUS_EMOJI_MAP: dict[str, str] = {
    "to-do": "\U0001f534 to-do",
    "in progress": "\U0001f7e0 in progress",
    "waiting": "\U0001f535 waiting",
    "done": "\U0001f7e2 done",
    "backlog": "\U0001f7e3 backlog",
    "cancelled": "\u26ab cancelled",
}

# Statuses that already have the emoji prefix
_VALID_STATUSES = frozenset(STATUS_EMOJI_MAP.values())

# Key ordering for frontmatter — these come first, in order; rest alphabetical
_FM_KEY_ORDER = ["type", "category", "created", "slug", "timestamp", "area"]

_YAML_SPECIAL_CHARS = re.compile(r'[:#"\'`\[\]{}|>&*!?,]')


# ── URL Normalization ──────────────────────────────────────────────────────

_TRACKING_PARAMS = frozenset(
    {
        "utm_source",
        "utm_medium",
        "utm_campaign",
        "utm_term",
        "utm_content",
        "ref",
        "fbclid",
        "gclid",
        "mc_cid",
        "mc_eid",
        "s",
        "si",
    }
)


def normalize_url(url: str) -> str:
    """Normalize URL for dedup: strip tracking params, trailing slash, www.

    YouTube canonicalization: youtu.be/X, shorts/X, embed/X → youtube.com/watch?v=X.
    """
    parsed = urlparse(url)
    hostname = parsed.hostname or ""
    if hostname.startswith("www."):
        hostname = hostname[4:]

    path = parsed.path
    query_params = parse_qs(parsed.query, keep_blank_values=False)
    if hostname in ("youtu.be",):
        video_id = path.lstrip("/").split("/")[0] if path else ""
        if video_id:
            hostname = "youtube.com"
            path = "/watch"
            query_params = {"v": [video_id]}
    elif hostname == "youtube.com":
        for prefix in ("/shorts/", "/embed/"):
            if path.startswith(prefix):
                video_id = path[len(prefix) :].split("/")[0].split("?")[0]
                path = "/watch"
                query_params = {"v": [video_id]}
                break

    clean_params = {
        k: v for k, v in query_params.items() if k.lower() not in _TRACKING_PARAMS
    }
    query_str = urlencode(clean_params, doseq=True)

    if path != "/" and path.endswith("/"):
        path = path.rstrip("/")

    return urlunparse((parsed.scheme, hostname, path, "", query_str, ""))


# ── Schema loading ──────────────────────────────────────────────────────────

_schema_loaded = False


def _try_load_schema() -> None:
    """One-shot: override constants from life-os-schema.yaml if available."""
    global _schema_loaded, VALID_TYPES, VALID_AREAS, STATUS_EMOJI_MAP, _VALID_STATUSES
    if _schema_loaded:
        return
    _schema_loaded = True

    try:
        import yaml
    except ImportError:
        return

    schema_path = contract_path("life-os-schema")
    if not schema_path.exists():
        return

    try:
        with open(schema_path, encoding="utf-8") as f:
            schema = yaml.safe_load(f)
    except Exception:
        return

    # Types
    types_section = schema.get("types")
    if isinstance(types_section, dict):
        VALID_TYPES = frozenset(types_section.keys())

    # Areas
    enums = schema.get("enums", {})
    area_list = enums.get("area")
    if isinstance(area_list, list):
        # Items may be strings or dicts; extract string values
        areas = set()
        for item in area_list:
            if isinstance(item, str):
                # Strip inline comments like "work      # Professional..."
                areas.add(item.split()[0] if item.strip() else item)
            elif isinstance(item, dict) and "value" in item:
                areas.add(item["value"])
        if areas:
            VALID_AREAS = frozenset(areas)

    # Statuses
    status_list = enums.get("status")
    if isinstance(status_list, list):
        new_map: dict[str, str] = {}
        for item in status_list:
            if isinstance(item, dict) and "value" in item:
                val = item["value"]
                # Extract the plain text after the emoji
                plain = re.sub(
                    r"^[\U0001f300-\U0001faff\u2600-\u27bf\u26ab\u2b50]+\s*", "", val
                ).strip()
                if plain:
                    new_map[plain] = val
        if new_map:
            STATUS_EMOJI_MAP = new_map
            _VALID_STATUSES = frozenset(new_map.values())


# ── Public API ──────────────────────────────────────────────────────────────


def escape_title(title: str) -> str:
    """Escape title for YAML frontmatter double-quoted scalar."""
    return (
        title.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\n", " ")
        .replace("\r", "")
        .replace("\t", " ")
        .replace("#", "`#`")
    )


def title_to_slug(title: str, max_len: int = 40) -> str:
    """Convert title to kebab-case slug. Lowercase, hyphens, no special chars.

    >>> title_to_slug("Hello World! This is a Test")
    'hello-world-this-is-test'
    """
    words = re.findall(r"[a-z0-9]+", title.lower())
    words = [w for w in words if len(w) >= 2]
    result = ""
    for w in words:
        candidate = f"{result}-{w}" if result else w
        if len(candidate) > max_len:
            break
        result = candidate
    return result or "note"


def yaml_list(items: list[str]) -> str:
    """Render a list as inline YAML array.

    >>> yaml_list(["video", "code"])
    '["video", "code"]'
    """
    if not items:
        return "[]"
    return "[" + ", ".join(f'"{i}"' for i in items) + "]"


def make_slug(
    dt: datetime,
    prefix: str,
    title: str,
    *,
    notes_dir: Path | None = None,
    max_len: int = 40,
) -> tuple[str, str]:
    """Generate YYYYMMDD-HHMM-prefix-title-words slug with collision detection.

    Args:
        dt: Datetime for the timestamp portion.
        prefix: Note type prefix (e.g. "bookmark", "mail", "task").
                Pass "" to omit the prefix segment.
        title: Human title to slugify.
        notes_dir: Directory to check for collisions. If None, no collision check.
        max_len: Max length for the title portion of the slug.

    Returns:
        (slug, timestamp_str) where timestamp_str is YYYYMMDD-HHMM.
    """
    ts = dt.strftime("%Y%m%d-%H%M")
    title_part = title_to_slug(title, max_len)
    if prefix:
        base = f"{ts}-{prefix}-{title_part}" if title_part else f"{ts}-{prefix}"
    else:
        base = f"{ts}-{title_part}" if title_part else f"{ts}-note"

    if notes_dir is None:
        return base, ts

    filepath = notes_dir / f"{base}.md"
    if not filepath.exists():
        return base, ts

    counter = 2
    while (notes_dir / f"{base}-{counter}.md").exists():
        counter += 1
    return f"{base}-{counter}", ts


def normalize_status(status: str | None) -> str | None:
    """Auto-correct bare status to emoji-prefixed format.

    >>> normalize_status("to-do")
    '🔴 to-do'
    >>> normalize_status("🔴 to-do")
    '🔴 to-do'
    >>> normalize_status(None) is None
    True
    """
    if status is None:
        return None
    _try_load_schema()
    # Already correct
    if status in _VALID_STATUSES:
        return status
    # Bare status → add emoji
    if status in STATUS_EMOJI_MAP:
        return STATUS_EMOJI_MAP[status]
    return status


def validate_frontmatter(fm: dict[str, Any]) -> list[str]:
    """Validate frontmatter dict against life-os-schema rules.

    Returns list of error strings (empty = valid).
    """
    _try_load_schema()
    errors: list[str] = []

    # Required fields
    for field in ("type", "category", "created", "slug", "timestamp"):
        if field not in fm or fm[field] is None or str(fm[field]).strip() == "":
            errors.append(f"Missing required field: {field}")

    # Type validation
    note_type = fm.get("type")
    if note_type and note_type not in VALID_TYPES:
        errors.append(f"Invalid type: '{note_type}' (valid: {sorted(VALID_TYPES)})")

    # Area validation
    area = fm.get("area")
    if area and area not in VALID_AREAS:
        errors.append(f"Invalid area: '{area}' (valid: {sorted(VALID_AREAS)})")

    # Status validation
    status = fm.get("status")
    if status is not None:
        normalized = normalize_status(status)
        if normalized not in _VALID_STATUSES and status not in STATUS_EMOJI_MAP:
            errors.append(
                f"Invalid status: '{status}' (valid: {sorted(_VALID_STATUSES)})"
            )

    # Created format
    created = fm.get("created")
    if created is not None and not isinstance(created, date):
        created_str = str(created)
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", created_str):
            errors.append(
                f"Invalid created format: '{created_str}' (expected YYYY-MM-DD)"
            )

    # Timestamp format
    timestamp = fm.get("timestamp")
    if timestamp is not None:
        ts_str = str(timestamp)
        if not re.match(r"^\d{8}-\d{4}$", ts_str):
            errors.append(
                f"Invalid timestamp format: '{ts_str}' (expected YYYYMMDD-HHmm)"
            )

    return errors


def _render_fm_value(value: Any) -> str:
    """Render a single frontmatter value as a YAML string."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, date) and not isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, list):
        return yaml_list([str(item) for item in value])
    s = str(value)
    # Quote strings with special YAML characters
    if _YAML_SPECIAL_CHARS.search(s) or s.startswith(("'", '"')):
        escaped = s.replace("\\", "\\\\").replace('"', '\\"')
        return f'"{escaped}"'
    return s


def render_frontmatter(
    fm: dict[str, Any],
    *,
    preserve_empty_keys: set[str] | None = None,
) -> str:
    """Render a frontmatter dict as YAML between --- fences.

    Key order: type, category, created, slug, timestamp, area, then alphabetical.
    None/empty values are omitted unless their key is listed in
    preserve_empty_keys.
    Status is auto-corrected to emoji format.
    """
    # Auto-correct status
    if "status" in fm:
        fm = dict(fm)  # shallow copy to avoid mutating caller's dict
        fm["status"] = normalize_status(fm["status"])

    lines = ["---"]

    # Ordered keys first
    seen = set()
    for key in _FM_KEY_ORDER:
        if key in fm and fm[key] is not None:
            rendered = _render_fm_value(fm[key])
            if rendered:
                lines.append(f"{key}: {rendered}")
            elif preserve_empty_keys and key in preserve_empty_keys:
                lines.append(f"{key}:")
            seen.add(key)

    # Remaining keys alphabetically
    for key in sorted(fm.keys()):
        if key in seen:
            continue
        val = fm[key]
        if val is None:
            continue
        rendered = _render_fm_value(val)
        if rendered:
            lines.append(f"{key}: {rendered}")
        elif preserve_empty_keys and key in preserve_empty_keys:
            lines.append(f"{key}:")

    lines.append("---")
    return "\n".join(lines)


def write_vault_note(
    filepath: Path,
    frontmatter: dict[str, Any],
    body: str,
    *,
    encoding: str = "utf-8",
    strict: bool = False,
    preserve_empty_keys: set[str] | None = None,
) -> Path:
    """Write a complete vault note (frontmatter + body) atomically.

    Args:
        filepath: Destination path for the note.
        frontmatter: Dict of frontmatter fields.
        body: Markdown body content.
        encoding: File encoding (default utf-8).
        strict: If True, raise ValueError on validation errors.
                If False (default), log warnings.

    Returns:
        The filepath that was written.

    Raises:
        ValueError: If strict=True and validation fails.
    """
    errors = validate_frontmatter(frontmatter)
    if errors:
        if strict:
            raise ValueError(f"Frontmatter validation failed: {'; '.join(errors)}")
        for err in errors:
            logger.warning("Frontmatter warning: %s (file: %s)", err, filepath.name)

    fm_text = render_frontmatter(frontmatter, preserve_empty_keys=preserve_empty_keys)

    # Ensure body has exactly one leading newline after frontmatter
    body_stripped = body.lstrip("\n")
    content = f"{fm_text}\n\n{body_stripped}"

    # Ensure trailing newline
    if not content.endswith("\n"):
        content += "\n"

    # Atomic write: write to temp file in same directory, then rename
    dirpath = filepath.parent
    dirpath.mkdir(parents=True, exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(dir=dirpath, suffix=".md.tmp")
    try:
        with os.fdopen(fd, "w", encoding=encoding) as f:
            f.write(content)
        os.replace(tmp_path, filepath)
    except Exception:
        # Clean up temp file on failure
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise

    return filepath


def link_in_daily(slug: str) -> None:
    """Link a note in the daily note via brain_lib. Swallows errors."""
    try:
        import brain_lib

        brain_lib.link_note_in_daily(slug)
    except Exception as exc:
        logger.warning("Failed to link %s in daily note: %s", slug, exc)
