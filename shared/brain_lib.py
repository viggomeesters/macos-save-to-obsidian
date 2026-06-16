"""brain_lib — central config + helpers for the Life OS toolchain.

Repo paths (centralized 2026-04-18 — replaces hardcoded ~/Dev/life-os-X strings).
Read from paths.json (~/.config/life-os/paths.json) "repos" section. Use:

    from brain_lib import cfg
    cfg.repo_pipeline                       # Path to ~/Dev/life-os-pipeline
    cfg.repo_root("meeting")                # Path to ~/Dev/life-os-meeting

Add new repos by extending paths.json.repos AND Config.__init__ section 8.
Future renames: only paths.json + brain_lib need updating; consumers inherit.
"""

from __future__ import annotations

import fcntl
import inspect
import json
import os
import random
import re
import sys
from pathlib import Path
from typing import Any, Optional
from datetime import datetime, date

try:
    from contract_paths import contract_path
except ImportError:  # pragma: no cover - package import fallback
    from .contract_paths import contract_path

try:
    import yaml as _yaml

    _HAS_YAML = True
except ImportError:
    _HAS_YAML = False

# ============================================
# BASE PATHS
# ============================================

ROOT = Path(__file__).resolve().parents[1]  # = ~/Dev/life-os-core
# Was parents[2] (~/Dev/agent-brain) — repo split 2026-04

# paths.json resolution: env var → well-known location → repo-relative (current behavior)
_paths_env = os.environ.get("LIFE_OS_PATHS")
_paths_wellknown = Path.home() / ".config" / "life-os" / "paths.json"
_paths_repo = ROOT / "context" / "paths.json"
PATHS_FILE = (
    Path(_paths_env)
    if _paths_env and Path(_paths_env).exists()
    else _paths_wellknown
    if _paths_wellknown.exists()
    else _paths_repo
)
PROJECTS_FILE = (
    PATHS_FILE.parent / "projects.json"
    if PATHS_FILE != _paths_repo
    else ROOT / "context" / "projects.json"
)

# Schema version is read dynamically from the vault contract source of truth.
SCHEMA_FILE = contract_path("life-os-schema")

# ============================================
# DOTENV LOADER (no external dependency)
# ============================================


def _load_dotenv(env_path: Path | None = None) -> None:
    """Load KEY=VALUE pairs from .env into os.environ. No-op if file missing.

    Does NOT override existing env vars (env takes precedence over .env file).
    Supports # comments, empty lines, optional quoting, and export prefix.
    """
    path = env_path or ROOT / ".env"
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[7:]
        if "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip().strip("'\"")
        if key and key not in os.environ:
            os.environ[key] = value


_load_dotenv()

# ============================================
# CONFIG MANAGEMENT
# ============================================


def current_agent_name() -> str:
    """Best-effort runtime agent detection across supported platforms."""
    explicit = os.environ.get("AGENT_NAME", "").strip()
    if explicit:
        return explicit
    if (
        os.environ.get("CODEX_THREAD_ID")
        or os.environ.get("CODEX_SHELL")
        or os.environ.get("CODEX_CI")
    ):
        return "codex"
    if os.environ.get("CLAUDE_SESSION_ID") or os.environ.get("CLAUDECODE"):
        return "claude-code"
    if os.environ.get("GEMINI_CLI") or os.environ.get("GEMINI_MODEL"):
        return "gemini"
    return "agent"


def agent_coauthor_trailer(agent_name: str | None = None) -> str:
    """Return a co-author trailer that matches the current agent platform."""
    name = (agent_name or current_agent_name()).strip().lower()
    if name == "claude-code":
        return "Co-Authored-By: Claude Opus 4.6 <noreply@anthropic.com>"
    if name == "codex":
        return "Co-Authored-By: OpenAI Codex <noreply@openai.com>"
    if name == "gemini":
        return "Co-Authored-By: Gemini CLI <noreply@google.com>"
    return "Co-Authored-By: AI Agent <noreply@example.invalid>"


def _claude_project_memory_path(project_root: Path | None = None) -> Path:
    """Return Claude's project MEMORY.md path for a repo root."""
    root = (project_root or ROOT).expanduser().resolve()
    key = "-" + "-".join(part for part in root.parts if part and part != "/")
    return Path.home() / ".claude" / "projects" / key / "memory" / "MEMORY.md"


def feedback_memory_paths(project_root: Path | None = None) -> list[Path]:
    """Return memory files that are useful for agent feedback lookups."""
    candidates = [
        ROOT / "context" / "agent-memory.md",
        _claude_project_memory_path(project_root),
    ]
    deduped: list[Path] = []
    seen: set[str] = set()
    for path in candidates:
        key = str(path)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(path)
    return deduped


class Config:
    """Centralized configuration management with environment overrides."""

    def __init__(self):
        # Load static paths
        self.static_paths = self._load_static_paths()

        # 1. Resolve Vault Root
        self.vault_root = self._resolve_path(
            "VAULT_ROOT", "vault_root", "~/Library/Mobile Documents/iCloud~md~obsidian/Documents/vault"
        )

        # 2. Derived Vault Paths (Drieluik V5)
        self.vault_notes = self.vault_root / "10_notes"
        self.vault_files = self.vault_root / "20_files"
        self.vault_attachments = self.vault_files
        self.vault_media = self.vault_root / "30_media"
        self.vault_inbox = self.vault_root / "00_inbox"
        self.vault_system = self.vault_root / "system"

        # 3. Derived System Subpaths
        self.vault_chores = self.vault_system / "chores"
        self.vault_entities = self.vault_system / "entities"
        self.vault_anniversaries = self.vault_system / "anniversaries"
        self.vault_projects = self.vault_system / "projects"
        self.vault_context = self.vault_system / "context"

        # 4. Health Config
        health = self.static_paths.get("health", {})
        self.medication_preventive = health.get("medication_preventive", "")

        # 5. State Paths (centralized from paths.json, defaults to /tmp/)
        sp = self.static_paths.get("state_paths", {})
        self.state_task_cache = Path(
            sp.get("task_cache", "/tmp/brain-vault-tasks-cache.json")
        )
        self.state_write_queue = Path(
            sp.get("write_queue", "/tmp/brain-vault-write-queue.json")
        )
        self.state_pending_commits_prefix = sp.get(
            "pending_commits_prefix", "/tmp/brain-vault-pending-commit-paths"
        )

        # 6. System Paths (macOS-specific, centralized)
        sys_paths = self.static_paths.get("system_paths", {})
        self.messages_db = Path(
            sys_paths.get("messages_db", "~/Library/Messages/chat.db")
        ).expanduser()
        self.launch_agents = Path(
            sys_paths.get("launch_agents", "~/Library/LaunchAgents")
        ).expanduser()

        # 7. Agent Metadata
        self.agent_name = current_agent_name()

        # 8. Repo Roots (centralized — replace hardcoded ~/Dev/life-os-X paths)
        repos = self.static_paths.get("repos", {})
        dev_root = Path(self.static_paths.get("dev_root", "~/Dev")).expanduser()
        self.repo_core = Path(repos.get("core", dev_root / "life-os-core")).expanduser()
        self.repo_pipeline = Path(
            repos.get("pipeline", dev_root / "life-os-pipeline")
        ).expanduser()
        self.repo_skills = Path(
            repos.get("skills", dev_root / "life-os-skills")
        ).expanduser()
        self.repo_meeting = Path(
            repos.get("meeting", dev_root / "life-os-meeting")
        ).expanduser()
        self.repo_consumption = Path(
            repos.get("consumption", dev_root / "life-os-consumption")
        ).expanduser()

    def repo_root(self, name: str) -> Path:
        """Get a repo root by short name: core, pipeline, skills, meeting, consumption."""
        attr = f"repo_{name}"
        if not hasattr(self, attr):
            raise ValueError(
                f"Unknown repo: {name}. Known: core, pipeline, skills, meeting, consumption"
            )
        return getattr(self, attr)

    def skill_repos(self) -> list[Path]:
        """Repos that contain a skills/ directory, in resolution order."""
        return [
            self.repo_skills,
            self.repo_pipeline,
            self.repo_meeting,
            self.repo_consumption,
        ]

    def script_repos(self) -> list[Path]:
        """All repos (including core) that may contain scripts/."""
        return [self.repo_core, *self.skill_repos()]

    @property
    def commands_dir(self) -> Path:
        """Canonical location of slash-command definitions."""
        return self.repo_skills / "commands"

    def hooked_repos(self) -> list[Path]:
        """Repos that get the shared pre-commit hook installed.

        Only repos where the pre-commit checks actually test something
        (core: validator + pytest; pipeline: 3 integration test scripts).
        Skills/meeting/consumption currently have no meaningful checks, so
        the hook would be a no-op there.
        """
        return [self.repo_core, self.repo_pipeline]

    def _load_static_paths(self) -> dict:
        if not PATHS_FILE.exists():
            return {}
        try:
            data = json.loads(PATHS_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
        # Validate against schema if jsonschema available — non-fatal warning only.
        # Catches typos (vault-root vs vault_root) and missing required keys.
        try:
            import jsonschema  # type: ignore

            schema_path = contract_path("paths-schema")
            if schema_path.exists():
                schema = json.loads(schema_path.read_text(encoding="utf-8"))
                jsonschema.validate(data, schema)
        except ImportError:
            pass  # jsonschema optional dep
        except Exception as e:
            print(f"⚠️  paths.json schema warning: {e}", file=sys.stderr)
        return data

    def _resolve_path(self, env_key: str, json_key: str, default: str) -> Path:
        val = os.environ.get(env_key) or self.static_paths.get(json_key) or default
        return Path(val).expanduser().resolve()

    @staticmethod
    def get_schema_version() -> str:
        """Read schema version from life-os-schema.yaml (single source of truth)."""
        if not SCHEMA_FILE.exists():
            return ""
        try:
            text = SCHEMA_FILE.read_text(encoding="utf-8")
            if _HAS_YAML:
                data = _yaml.safe_load(text)
                if isinstance(data, dict):
                    return str(data.get("version", ""))
            # Fallback: regex for version field
            m = re.match(
                r'.*^version:\s*["\']?(\d+\.\d+)', text, re.MULTILINE | re.DOTALL
            )
            return m.group(1) if m else ""
        except Exception:
            return ""

    _REQUIRED_KEYS = ("vault_root", "brain_root", "vault_structure")

    def validate(self):
        """Verify config integrity: required keys, paths, schema version."""
        ok = True

        # 1. Required keys in paths.json
        for key in self._REQUIRED_KEYS:
            if key not in self.static_paths:
                print(f"⚠️  paths.json missing required key: {key}", file=sys.stderr)
                ok = False

        # 2. Vault root exists
        if not self.vault_root.exists():
            print(
                f"⚠️  VAULT_ROOT does not exist at {self.vault_root}",
                file=sys.stderr,
            )
            ok = False

        # 3. Schema version consistency (paths.json vs YAML source of truth)
        paths_version = self.static_paths.get("vault_schema_version", "")
        yaml_version = self.get_schema_version()
        if paths_version and yaml_version and paths_version != yaml_version:
            print(
                f"⚠️  Schema version drift: paths.json says {paths_version}, "
                f"life-os-schema.yaml says {yaml_version}",
                file=sys.stderr,
            )

        return ok


TIMELINE_NOTE_TYPES = frozenset(
    {"interaction", "purchase", "health", "entry", "task", "reference"}
)
STATIC_NOTE_DIR_ATTRS = {
    "entity": "vault_entities",
    "project": "vault_projects",
    "chore": "vault_chores",
    "context": "vault_context",
    "anniversary": "vault_anniversaries",
}
MEDIA_EXTENSIONS = frozenset(
    {
        ".jpg",
        ".jpeg",
        ".png",
        ".gif",
        ".webp",
        ".svg",
        ".heic",
        ".heif",
        ".tiff",
        ".tif",
        ".bmp",
        ".mp3",
        ".m4a",
        ".wav",
        ".aac",
        ".flac",
        ".ogg",
        ".mp4",
        ".mov",
        ".m4v",
        ".webm",
        ".avi",
        ".mkv",
    }
)
MEDIA_MIME_PREFIXES = ("image/", "audio/", "video/")


def _coerce_datetime(value: str | datetime | date | None = None) -> datetime:
    """Return a datetime for vault path routing."""
    if value is None:
        return datetime.now()
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)

    text = str(value).strip()
    candidates: list[tuple[str, str]] = []
    if re.match(r"^\d{8}-\d{4}", text):
        candidates.append((text[:13], "%Y%m%d-%H%M"))
    if re.match(r"^\d{8}", text):
        candidates.append((text[:8], "%Y%m%d"))
    if re.match(r"^\d{4}-\d{2}-\d{2}", text):
        candidates.append((text[:10], "%Y-%m-%d"))
    for candidate, fmt in candidates:
        try:
            return datetime.strptime(candidate, fmt)
        except ValueError:
            continue
    raise ValueError(f"cannot derive vault month from timestamp: {value!r}")


def vault_month(value: str | datetime | date | None = None) -> str:
    """Return `YYYY-MM` for a vault timestamp/date."""
    return _coerce_datetime(value).strftime("%Y-%m")


def canonical_note_dir(
    note_type: str,
    timestamp: str | datetime | date | None = None,
    *,
    create: bool = False,
) -> Path:
    """Return the canonical directory for a Life OS note type."""
    if note_type in TIMELINE_NOTE_TYPES:
        path = cfg.vault_notes / vault_month(timestamp)
    elif note_type in STATIC_NOTE_DIR_ATTRS:
        path = getattr(cfg, STATIC_NOTE_DIR_ATTRS[note_type])
    else:
        raise ValueError(f"unknown canonical note type: {note_type!r}")
    if create:
        path.mkdir(parents=True, exist_ok=True)
    return path


def canonical_note_path(
    note_type: str,
    slug: str,
    timestamp: str | datetime | date | None = None,
    *,
    create_parent: bool = False,
) -> Path:
    """Return the canonical markdown path for a note slug and type."""
    return canonical_note_dir(note_type, timestamp, create=create_parent) / f"{slug}.md"


def artifact_dir(
    filename_or_mime: str,
    timestamp: str | datetime | date | None = None,
    *,
    create: bool = False,
) -> Path:
    """Return `30_media/YYYY-MM` for media, otherwise `20_files/YYYY-MM`."""
    value = (filename_or_mime or "").strip().lower()
    suffix = Path(value.split(";", 1)[0]).suffix
    is_media = suffix in MEDIA_EXTENSIONS or value.startswith(MEDIA_MIME_PREFIXES)
    base = cfg.vault_media if is_media else cfg.vault_files
    path = base / vault_month(timestamp)
    if create:
        path.mkdir(parents=True, exist_ok=True)
    return path


def canonical_artifact_path(
    filename: str,
    timestamp: str | datetime | date | None = None,
    *,
    mime_type: str | None = None,
    create_parent: bool = False,
) -> Path:
    """Return the canonical path for an attachment or media file."""
    route_key = mime_type or filename
    return artifact_dir(route_key, timestamp, create=create_parent) / filename


def vault_relative_path(path: Path, *, strip_markdown_suffix: bool = False) -> str:
    """Return a vault-relative POSIX path for Obsidian/Raycast output."""
    rel = path.resolve().relative_to(cfg.vault_root.resolve())
    if strip_markdown_suffix and rel.suffix == ".md":
        rel = rel.with_suffix("")
    return rel.as_posix()


def obsidian_file_ref(path: Path) -> str:
    """Return the Obsidian `file=` value for a vault path."""
    return vault_relative_path(path, strip_markdown_suffix=True)


# Singleton instance
cfg = Config()

# Module-level aliases for backward compat (scripts reference brain_lib.STATE_DIR)
STATE_DIR = Path(cfg.static_paths.get("state_dir", str(ROOT / "context" / "state")))

# ============================================
# TASK ID GENERATION
# ============================================


def new_task_id() -> str:
    """Generate a unique task ID: tsk-YYYYMMDDHHmmss-<hex4>."""
    return (
        f"tsk-{datetime.now().strftime('%Y%m%d%H%M%S')}-{random.randint(0, 0xFFFF):04x}"
    )


# ============================================
# LEGACY HELPERS (deprecated — use cfg.* directly)
# ============================================


def load_paths() -> dict[str, str]:
    """Deprecated: use cfg.static_paths instead."""
    return cfg.static_paths


def get_vault_root() -> Path:
    """Deprecated: use cfg.vault_root instead."""
    return cfg.vault_root


# ============================================
# LOGGING
# ============================================

import logging


def get_logger(name: str, *, verbose: bool | None = None) -> logging.Logger:
    """Return a logger with format ``[name] message``, writing to stderr."""
    logger = logging.getLogger(f"brain.{name}")
    if logger.handlers:
        return logger
    if verbose is None:
        verbose = "--verbose" in sys.argv or "-v" in sys.argv
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(logging.Formatter(f"[{name}] %(message)s"))
    logger.addHandler(handler)
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    logger.propagate = False
    return logger


# ============================================
# JSON I/O
# ============================================


def read_json(path: Path, default: Any = None) -> Any:
    if not path.exists():
        return default if default is not None else {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default if default is not None else {}


def write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    out = json.dumps(data, indent=2, ensure_ascii=False) + "\n"
    path.write_text(out, encoding="utf-8")


# ============================================
# VERSIONED STATE I/O
# ============================================

# Central TTL registry — all ephemeral state TTLs in one place
STATE_TTLS = {
    "task_cache": 120,  # 2 min — cross-process task scan cache
    "pending_commits": 6 * 3600,  # 6 hrs — bridges CLI turns, not day-old state
    "write_queue": None,  # persistent — no TTL, replayed at startup
}


def read_versioned_state(path: Path, expected_version: int, default: Any = None) -> Any:
    """Read a JSON state file and check its _meta.version field.

    Returns the data (without _meta) if version matches.
    Returns default and logs a warning if version mismatches.
    """
    data = read_json(path, default)
    if data is default or data is None:
        return default
    if not isinstance(data, dict):
        return data
    meta = data.get("_meta", {})
    found_version = meta.get("version", 0)
    if found_version != expected_version:
        print(
            f"⚠️  State version mismatch in {path.name}: "
            f"expected v{expected_version}, found v{found_version}. "
            f"State will be reset.",
            file=sys.stderr,
        )
        return default
    # Strip _meta from returned data
    return {k: v for k, v in data.items() if k != "_meta"}


def write_versioned_state(path: Path, data: dict, version: int) -> None:
    """Write a JSON state file with _meta.version for migration safety."""
    stamped = {"_meta": {"version": version}, **data}
    write_json(path, stamped)


# ============================================
# DATE UTILITIES
# ============================================

_DATE_FORMATS = ("%Y-%m-%d", "%Y%m%d", "%d %b %Y")

# ============================================
# VAULT / MARKDOWN
# ============================================


def _parse_frontmatter_custom(fm_text: str) -> dict[str, Any]:
    """Custom (legacy) frontmatter parser — handles simple key: value pairs."""
    fm: dict[str, Any] = {}
    current_key: str | None = None
    current_list: list | None = None
    quotes = chr(34) + chr(39)
    for line in fm_text.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        if stripped.startswith("- ") and current_key:
            if current_list is None:
                current_list = []
            val = stripped[2:].strip().strip(quotes)
            current_list.append(val)
            fm[current_key] = current_list
            continue
        if ":" in stripped:
            parts = stripped.split(":", 1)
            k = parts[0].strip()
            v = parts[1].strip().strip(quotes + " ")
            if v.startswith("[") and v.endswith("]"):
                items = v[1:-1].split(",")
                fm[k] = [item.strip().strip(quotes) for item in items if item.strip()]
                current_key = None
                current_list = None
            elif not v:
                current_key = k
                current_list = []
                fm[k] = current_list
            else:
                fm[k] = v
                current_key = k
                current_list = None
    return fm


def _normalize_yaml_value(val: Any) -> Any:
    """Coerce pyyaml-parsed values to strings (preserving lists of strings).

    The vault tooling expects all scalar values as strings — pyyaml converts
    booleans, ints, floats, and dates to native Python types which would break
    downstream comparisons.
    """
    if val is None:
        return ""
    if isinstance(val, bool):
        return "true" if val else "false"
    if isinstance(val, list):
        return [str(item) for item in val]
    return str(val)


def _find_closing_delimiter(content: str, start: int = 3) -> int:
    """Find the closing --- delimiter that starts at the beginning of a line.

    Returns the index of the closing ---, or -1 if not found.
    Skips --- that appears mid-line (e.g. inside quoted YAML values).
    """
    search_from = start
    while True:
        idx = content.find("---", search_from)
        if idx == -1:
            return -1
        if idx == 0 or content[idx - 1] == "\n":
            return idx
        search_from = idx + 3
    return -1


def parse_frontmatter(content: str) -> dict[str, Any]:
    if not content.startswith("---"):
        return {}
    end = _find_closing_delimiter(content)
    if end == -1:
        return {}
    fm_text = content[3:end].strip()
    if not fm_text:
        return {}

    # Primary: pyyaml (handles multi-line, nested, anchors, etc.)
    if _HAS_YAML:
        try:
            raw = _yaml.safe_load(fm_text)
            if isinstance(raw, dict):
                return {str(k): _normalize_yaml_value(v) for k, v in raw.items()}
        except _yaml.YAMLError:
            pass  # fall through to custom parser

    # Fallback: custom parser (handles edge cases pyyaml rejects)
    return _parse_frontmatter_custom(fm_text)


def read_note(path: Path) -> tuple[dict[str, Any], str]:
    try:
        content = path.read_text(encoding="utf-8")
        return parse_frontmatter(content), content
    except Exception:
        return {}, ""


def find_daily_note(dt: Optional[datetime] = None) -> Optional[Path]:
    """Find today's daily note. Convention: YYYYMMDD-0600-daily.md."""
    d = dt or datetime.now()
    today = d.strftime("%Y%m%d")
    year_month = d.strftime("%Y-%m")

    # Check chronologische folder eerst
    canonical = cfg.vault_notes / year_month / f"{today}-0600-daily.md"
    if canonical.exists():
        return canonical

    # Fallback: recursief zoeken in 10_notes
    for root, dirs, files in os.walk(cfg.vault_notes):
        dirs.sort()
        for f in sorted(files):
            if f.startswith(today) and f.endswith("-daily.md"):
                return Path(root) / f
    return None


_NL_DAYS = {
    "Monday": "Maandag",
    "Tuesday": "Dinsdag",
    "Wednesday": "Woensdag",
    "Thursday": "Donderdag",
    "Friday": "Vrijdag",
    "Saturday": "Zaterdag",
    "Sunday": "Zondag",
}
_NL_MONTHS = {
    1: "januari",
    2: "februari",
    3: "maart",
    4: "april",
    5: "mei",
    6: "juni",
    7: "juli",
    8: "augustus",
    9: "september",
    10: "oktober",
    11: "november",
    12: "december",
}


def _nl_date(dt: datetime) -> str:
    """Format datetime as 'Woensdag 11 maart 2026'."""
    day_name = _NL_DAYS[dt.strftime("%A")]
    return f"{day_name} {dt.day} {_NL_MONTHS[dt.month]} {dt.year}"


def ensure_daily_note() -> Optional[Path]:
    """Ensure today's daily note exists. Creates one if missing. Returns path."""
    now = datetime.now()
    today = now.strftime("%Y%m%d")
    year_month = now.strftime("%Y-%m")
    try:
        found = find_daily_note(now)
        if found:
            return found
        # Create new daily note with canonical 0600 timestamp in correct folder
        filename = f"{today}-0600-daily.md"
        note_dir = cfg.vault_notes / year_month
        note_dir.mkdir(parents=True, exist_ok=True)
        path = note_dir / filename
        slug = f"{today}-0600-daily"
        content = f"""---
type: entry
category: daily
created: {now.strftime("%Y-%m-%d")}
slug: {slug}
timestamp: {today}-0600
area: self
title: "Daily Note"
topics: [daily]
---

# {_nl_date(now)}

- [Start Of Day](raycast://script-commands/start-of-day).

# Viggo


# Automated

## Log
"""
        path.write_text(content, encoding="utf-8")
        return path
    except Exception as e:
        print(f"Could not create daily note: {e}", file=sys.stderr)
        return None


def detect_project(cwd: str | Path | None = None) -> dict | None:
    """Match a working directory to a project from projects.json.

    Returns the matching project dict (id, path, vault_note, type) or None.
    """
    projects = read_json(PROJECTS_FILE, {})
    active = projects.get("active", [])
    if not active:
        return None

    if cwd is None:
        cwd = Path.cwd()
    cwd = Path(cwd).expanduser().resolve()
    cwd_str = str(cwd)

    for p in active:
        project_path = Path(p.get("path", "")).expanduser().resolve()
        if cwd_str == str(project_path) or cwd_str.startswith(str(project_path) + "/"):
            return p
    return None


_project_code_cache: dict[str, str] = {}


def resolve_project_code(project_slug: str) -> str:
    """Resolve a project slug to its short code from vault project note.

    Reads the `code` field from the project note frontmatter.
    Returns the code if found, otherwise returns the slug unchanged.
    Results are cached in-memory for the process lifetime.
    """
    if not project_slug:
        return ""
    if project_slug in _project_code_cache:
        return _project_code_cache[project_slug]
    projects_dir = cfg.vault_projects
    if not projects_dir.exists():
        return project_slug
    for f in projects_dir.glob("*.md"):
        fm, _ = read_note(f)
        if fm.get("slug") == project_slug or f.stem == project_slug:
            code = fm.get("code", "")
            result = str(code).lower() if code else project_slug
            _project_code_cache[project_slug] = result
            return result
    _project_code_cache[project_slug] = project_slug
    return project_slug


def parse_date(value: Any) -> Optional[date]:
    """Parse a date from frontmatter or slug. Returns None on failure.

    Handles: '2026-03-07', '20260307', '07 Mar 2026', datetime objects,
    quoted strings, and truncates timestamps to date-only.
    """
    if value is None:
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()
    s = str(value).strip().strip("'\"")
    if not s:
        return None
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def toggle_sweep_checkbox(slug: str) -> None:
    """Toggle a matching - [ ] checkbox to - [x] in today's daily note sweep."""
    daily_path = find_daily_note()
    if not daily_path:
        return
    try:
        content = daily_path.read_text(encoding="utf-8")
    except (FileNotFoundError, PermissionError):
        return
    pattern = re.compile(
        r"^(- \[) \](.+\[\[" + re.escape(slug) + r"(?:\|[^\]]*)?]])",
        re.MULTILINE,
    )
    new_content, n = pattern.subn(r"\1x]\2", content)
    if n > 0:
        daily_path.write_text(new_content, encoding="utf-8")


# Patterns that YAML would interpret as non-string types
_YAML_AMBIGUOUS = re.compile(
    r"^("
    r"true|false|yes|no|on|off"  # booleans
    r"|null|~"  # null
    r"|\d+\.?\d*([eE][+-]?\d+)?"  # integers/floats (123, 1.5, 1e10)
    r"|\.?\d+([eE][+-]?\d+)?"  # .5-style floats
    r"|0[xXoObB][\da-fA-F]+"  # hex/octal/binary
    r"|[+-]?(\.inf|\.nan)"  # special floats
    r"|\d{4}-\d{2}-\d{2}"  # dates (2026-04-04)
    r")$",
    re.IGNORECASE,
)


_YAML_DATE_FIELDS = {
    "birthday",
    "created",
    "date",
    "due",
    "last_done",
    "snoozed_until",
}


def _yaml_format_value(val: Any, key: str | None = None) -> str:
    """Format a value for YAML frontmatter, quoting strings that YAML would misinterpret."""
    if val is None or val == "":
        return ""
    if isinstance(val, list):
        return f"[{', '.join(_yaml_format_value(v) for v in val)}]"
    s = str(val)
    if key in _YAML_DATE_FIELDS and re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return s
    # Quote strings that look like YAML types, contain special chars, or are empty
    if (
        _YAML_AMBIGUOUS.match(s)
        or s != s.strip()  # leading/trailing whitespace (incl. whitespace-only)
        or s.startswith(("{", "[", "'", '"', "&", "*", "!", "|", ">", "%", "@", "#"))
        or " #" in s  # inline YAML comments
        or ": " in s
        or s.endswith(":")
        or "\n" in s
        or s.startswith("---")
    ):
        # Use single quotes, escape embedded single quotes
        return f"'{s.replace(chr(39), chr(39) + chr(39))}'"
    return s


def update_frontmatter(file_path: Path, updates: dict[str, Any]) -> bool:
    """Merge updates into existing YAML frontmatter, preserve body."""
    if not file_path.exists():
        return False
    content = file_path.read_text(encoding="utf-8")
    if not content.startswith("---"):
        return False
    end_fm = _find_closing_delimiter(content)
    if end_fm == -1:
        return False
    fm_text = content[3:end_fm].strip()
    fm_lines = fm_text.splitlines()
    body = content[end_fm + 3 :]

    new_fm_lines = []
    processed_keys: set[str] = set()
    skip_list_items = False

    for line in fm_lines:
        stripped = line.strip()
        # Skip continuation list items for a key we're replacing
        if skip_list_items:
            if stripped.startswith("- "):
                continue
            skip_list_items = False

        if ":" in line and not stripped.startswith("- "):
            key = line.split(":", 1)[0].strip()
            if key in updates:
                val = updates[key]
                new_fm_lines.append(f"{key}: {_yaml_format_value(val, key)}")
                processed_keys.add(key)
                # Check if next lines are list items for this key
                skip_list_items = True
                continue
        new_fm_lines.append(line)

    # Append new keys not already in frontmatter
    for key, val in updates.items():
        if key not in processed_keys:
            new_fm_lines.append(f"{key}: {_yaml_format_value(val, key)}")

    new_content = "---\n" + "\n".join(new_fm_lines) + "\n---" + body
    file_path.write_text(new_content, encoding="utf-8")
    return True


def format_status_header(status: str, dt: datetime | None = None) -> str:
    """Format a ## status header line for vault notes.

    Args:
        status: Status with emoji, e.g. "🔴 to-do" or "🟢 done".
        dt: Timestamp for the header. Defaults to now.

    Returns:
        e.g. "## 🔴 to-do - 14 Apr 2026 at 10:13"
    """
    if dt is None:
        dt = datetime.now()
    date_str = dt.strftime("%-d %b %Y at %H:%M")
    return f"## {status} - {date_str}"


def format_source_quote(source: str, dt: datetime | None = None) -> str:
    """Format a > Bron: quote line for vault notes.

    Args:
        source: Source description, e.g. "iMessage self-messaging" or "Raycast Script Command".
        dt: Timestamp for the quote. Defaults to now.

    Returns:
        e.g. "> Bron: iMessage self-messaging op 14 april 2026 om 10:13"
    """
    if dt is None:
        dt = datetime.now()
    date_str = dt.strftime("%-d %B %Y om %H:%M")
    return f"> Bron: {source} op {date_str}"


def append_status_header(file_path: Path, status: str, note: str = "") -> bool:
    """Append a ## status header to vault task note body.

    Example output:
        ## 🟢 done - 27 Feb 2026 at 09:52
        Afgerond via agent ledger
    """
    if not file_path.exists():
        return False
    content = file_path.read_text(encoding="utf-8")
    header = "\n\n" + format_status_header(status)
    if note:
        header += f"\n{note}"
    content = content.rstrip() + header + "\n"
    file_path.write_text(content, encoding="utf-8")
    return True


def _insert_under_log(content: str, entry: str) -> str:
    """Insert an entry under ## Log in daily note content and sort chronologically.

    Args:
        content: Full daily note content
        entry: Line to insert (e.g. "- 08:15 - [[slug]]")

    Returns:
        Updated content with entry inserted and log sorted.
    """
    content = re.sub(r"^# Agent\b", "# Automated", content, count=1, flags=re.MULTILINE)
    log_match = re.search(r"^## Log\b", content, re.MULTILINE)
    if log_match:
        pos = content.find("\n", log_match.end())
        if pos == -1:
            content += f"\n{entry}\n"
        else:
            content = content[: pos + 1] + entry + "\n" + content[pos + 1 :]
    else:
        automated_match = re.search(r"^# Automated\b", content, re.MULTILINE)
        if automated_match:
            insert_pos = content.find("\n", automated_match.end())
            if insert_pos == -1:
                content += f"\n\n## Log\n{entry}\n"
            else:
                content = (
                    content[: insert_pos + 1]
                    + f"\n## Log\n{entry}\n"
                    + content[insert_pos + 1 :]
                )
        else:
            content = content.rstrip() + f"\n\n# Automated\n\n## Log\n{entry}\n"
    return _sort_daily_log(content)


def _humanize_script_name(name: str) -> str:
    stem = Path(name).stem
    stem = re.sub(r"^(raycast_|process_)", "", stem)
    words = re.split(r"[-_]+", stem)
    return " ".join(word.capitalize() for word in words if word) or "Vault"


def _infer_daily_log_source() -> str:
    for frame in inspect.stack()[2:]:
        filename = Path(frame.filename).name
        if filename in {
            "brain_lib.py",
            "vault_note_writer.py",
            "db_indexer.py",
            "<stdin>",
        }:
            continue
        return _humanize_script_name(filename)
    return "Vault"


def _format_daily_log_entry(
    ts_display: str,
    message: str,
    source: str | None = None,
    action: str | None = None,
) -> str:
    source_label = source or _infer_daily_log_source()
    action_label = (action or "UPDATED").upper()
    return f"- {ts_display} - {source_label} - {action_label} - {message}"


def log_vault(msg: str, source: str | None = None, action: str = "UPDATED") -> None:
    """Log a message to today's daily note under ## Log.

    Messages should contain a [[wikilink]] to maintain the atomic notes principle.
    Plain text without links will trigger a stderr warning.
    """
    if "[[" not in msg:
        print(
            f"log_vault warning: message has no [[wikilink]]: {msg[:60]}",
            file=sys.stderr,
        )
    ts = datetime.now().strftime("%H:%M")
    try:
        daily_note = ensure_daily_note()
        if not daily_note:
            return

        fd = os.open(str(daily_note), os.O_RDWR)
        try:
            fcntl.flock(fd, fcntl.LOCK_EX)
            content = daily_note.read_text(encoding="utf-8")
            content = _insert_under_log(
                content,
                _format_daily_log_entry(ts, msg, source=source, action=action),
            )
            daily_note.write_text(content, encoding="utf-8")
        finally:
            fcntl.flock(fd, fcntl.LOCK_UN)
            os.close(fd)
    except Exception as e:
        print(f"Logging to vault failed: {e}", file=sys.stderr)


def _sort_daily_log(content: str) -> str:
    """Sort list items in ## Log section by their HH:MM prefix.

    Collects ALL log entries (even across blank lines) between ## Log and
    the next heading or EOF, sorts them chronologically, and emits a single
    contiguous block — no stray blank lines. Non-entry lines (plain text
    without HH:MM prefix) are preserved but moved after sorted entries.
    """
    log_match = re.search(r"^## Log\n", content, re.MULTILINE)
    if not log_match:
        return content
    after_header = log_match.end()
    # Find next heading or EOF to delimit the log section
    next_heading = re.search(r"^## ", content[after_header:], re.MULTILINE)
    section_end = after_header + next_heading.start() if next_heading else len(content)
    section = content[after_header:section_end]
    entry_re = re.compile(r"^- (?P<time>\d{2}:\d{2})(?: -)? .+", re.MULTILINE)
    entries = list(entry_re.finditer(section))
    non_entries = [
        l for l in section.splitlines() if l.strip() and not entry_re.match(l)
    ]
    if not entries:
        return content
    sorted_entries = sorted(entries, key=lambda m: m.group("time"))
    sorted_entries = [m.group(0) for m in sorted_entries]
    sorted_block = "\n".join(sorted_entries) + "\n"
    if non_entries:
        sorted_block += "\n".join(non_entries) + "\n"
    return content[:after_header] + sorted_block + content[section_end:]


def link_note_in_daily(
    slug: str,
    timestamp: Optional[str] = None,
    source: str | None = None,
    action: str = "CREATED",
) -> None:
    """Add [[slug]] to ## Log in today's daily note. Idempotent.

    Args:
        timestamp: Optional YYYYMMDD-HHMM string. If given, the HH:MM is
                   extracted and shown instead of current time.
        source: Human-readable script/source name.
        action: CRUD action label, e.g. CREATED or UPDATED.
    """
    daily = ensure_daily_note()
    if not daily:
        return

    fd = os.open(str(daily), os.O_RDWR)
    try:
        fcntl.flock(fd, fcntl.LOCK_EX)
        content = daily.read_text(encoding="utf-8")

        # Skip if link already exists
        if f"[[{slug}" in content:
            return

        # Extract HH:MM from YYYYMMDD-HHMM timestamp, or fall back to slug
        ts_display = None
        for src in (timestamp, slug):
            if src:
                m = re.match(r"\d{8}-(\d{2})(\d{2})", src)
                if m:
                    ts_display = f"{m.group(1)}:{m.group(2)}"
                    break
        if not ts_display:
            ts_display = datetime.now().strftime("%H:%M")

        link = _format_daily_log_entry(
            ts_display,
            f"[[{slug}]]",
            source=source,
            action=action,
        )

        content = _insert_under_log(content, link)
        daily.write_text(content, encoding="utf-8")
    finally:
        fcntl.flock(fd, fcntl.LOCK_UN)
        os.close(fd)
