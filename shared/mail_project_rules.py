"""Fast project detection for saved mail.

Project metadata is the primary source. The editable mail-specific mapping lives
in the vault contract, while this module compiles a small JSON cache for the
Save Mail hot path.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

try:
    import yaml
except ImportError:  # pragma: no cover - bootstrap fallback
    yaml = None

try:
    import brain_lib
    from contract_paths import contract_path
except ImportError:  # pragma: no cover - package import fallback
    from . import brain_lib
    from .contract_paths import contract_path


CACHE_VERSION = 1
DEFAULT_CONFIDENCE = {
    "address": 1.0,
    "domain": 0.95,
    "project_code": 0.85,
    "mail_code": 0.9,
    "subject_rule": 0.75,
}
_MEMORY_CACHE: dict[str, dict[str, Any]] = {}


@dataclass(frozen=True)
class MailProjectMatch:
    project: str
    area: str
    source: str
    confidence: float
    matched: str
    candidates: tuple[str, ...] = ()

    @property
    def is_ambiguous(self) -> bool:
        return not self.project and bool(self.candidates)


def normalize_token(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip().casefold())


def _listify(value: Any) -> list[str]:
    if value is None or value == "":
        return []
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    raw = str(value).strip()
    if not raw:
        return []
    return [item.strip() for item in raw.split(",") if item.strip()]


def _read_rules(path: Path) -> dict[str, Any]:
    if not path.exists() or yaml is None:
        return {}
    try:
        data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _project_alias(slug: str) -> str:
    return re.sub(r"^\d{4}-\d{2}-", "", slug).strip("-")


def _project_files(projects_dir: Path) -> dict[str, float]:
    if not projects_dir.exists():
        return {}
    return {
        str(path): path.stat().st_mtime
        for path in sorted(projects_dir.glob("*.md"))
        if path.is_file()
    }


def _rules_mtime(path: Path) -> float:
    return path.stat().st_mtime if path.exists() else 0.0


def _cache_is_fresh(cache: dict[str, Any], rules_path: Path, projects_dir: Path) -> bool:
    if cache.get("version") != CACHE_VERSION:
        return False
    if cache.get("rules_path") != str(rules_path):
        return False
    if float(cache.get("rules_mtime") or 0.0) != _rules_mtime(rules_path):
        return False
    return cache.get("project_files") == _project_files(projects_dir)


def _load_cache(cache_path: Path, rules_path: Path, projects_dir: Path) -> Optional[dict[str, Any]]:
    try:
        cache = json.loads(cache_path.read_text(encoding="utf-8"))
    except Exception:
        return None
    if not isinstance(cache, dict):
        return None
    return cache if _cache_is_fresh(cache, rules_path, projects_dir) else None


def _write_cache(cache_path: Path, cache: dict[str, Any]) -> None:
    try:
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(
            json.dumps(cache, ensure_ascii=False, indent=2, sort_keys=True),
            encoding="utf-8",
        )
    except Exception:
        pass


def _iter_project_entries(projects_dir: Path) -> list[dict[str, Any]]:
    projects: list[dict[str, Any]] = []
    if not projects_dir.exists():
        return projects

    for path in sorted(projects_dir.glob("*.md")):
        fm, _ = brain_lib.read_note(path)
        if fm.get("type") != "project":
            continue
        slug = str(fm.get("slug") or path.stem).strip()
        if not slug:
            continue
        projects.append(
            {
                "slug": slug,
                "title": str(fm.get("title") or slug).strip(),
                "area": str(fm.get("area") or "").strip(),
                "code": str(fm.get("code") or "").strip(),
                "mail_codes": _listify(fm.get("mail_codes")),
                "path": str(path),
            }
        )
    return projects


def _code_entry(project: dict[str, Any], code: str, source: str) -> dict[str, Any]:
    return {
        "project": project["slug"],
        "slug": project["slug"],
        "area": project.get("area") or "",
        "code": code,
        "source": source,
        "title": project.get("title") or project["slug"],
        "path": project.get("path") or "",
        "mail_codes": project.get("mail_codes") or [],
    }


def _add_code(
    codes: dict[str, list[dict[str, Any]]],
    project: dict[str, Any],
    code: str,
    source: str,
) -> None:
    normalized = normalize_token(code)
    if len(normalized) < 2:
        return
    codes.setdefault(normalized, []).append(_code_entry(project, normalized, source))


def build_project_index(
    *,
    rules_path: Optional[Path] = None,
    projects_dir: Optional[Path] = None,
) -> dict[str, Any]:
    source_rules = rules_path or contract_path("mail-project-rules")
    source_projects = projects_dir or brain_lib.cfg.vault_projects
    rules = _read_rules(source_rules)
    codes: dict[str, list[dict[str, Any]]] = {}
    projects = _iter_project_entries(source_projects)

    for project in projects:
        if project.get("code"):
            _add_code(codes, project, str(project["code"]), "project_code")
        for mail_code in project.get("mail_codes") or []:
            _add_code(codes, project, mail_code, "mail_code")

    unique_codes: dict[str, dict[str, Any]] = {}
    conflicts: dict[str, list[dict[str, Any]]] = {}
    for code, entries in codes.items():
        project_slugs = {entry["project"] for entry in entries}
        if len(project_slugs) == 1:
            entry = dict(entries[0])
            entry["projects"] = sorted(project_slugs)
            unique_codes[code] = entry
        else:
            conflicts[code] = entries

    return {
        "version": CACHE_VERSION,
        "built_at": datetime.now().isoformat(timespec="seconds"),
        "rules_path": str(source_rules),
        "rules_mtime": _rules_mtime(source_rules),
        "projects_dir": str(source_projects),
        "project_files": _project_files(source_projects),
        "addresses": {
            normalize_token(key): value
            for key, value in (rules.get("addresses") or {}).items()
            if isinstance(value, dict)
        },
        "domains": {
            normalize_token(key): value
            for key, value in (rules.get("domains") or {}).items()
            if isinstance(value, dict)
        },
        "subject_rules": [
            rule
            for rule in (rules.get("subject_rules") or [])
            if isinstance(rule, dict) and rule.get("project")
        ],
        "blocked_domains": [
            normalize_token(domain) for domain in _listify(rules.get("blocked_domains"))
        ],
        "projects": projects,
        "codes": unique_codes,
        "conflicts": conflicts,
    }


def load_project_index(
    *,
    rules_path: Optional[Path] = None,
    projects_dir: Optional[Path] = None,
    cache_path: Optional[Path] = None,
    force_refresh: bool = False,
) -> dict[str, Any]:
    source_rules = rules_path or contract_path("mail-project-rules")
    source_projects = projects_dir or brain_lib.cfg.vault_projects
    source_cache = cache_path or (brain_lib.ROOT / ".brain" / "cache" / "mail-project-index.json")
    cache_key = f"{source_rules}|{source_projects}|{source_cache}"

    if not force_refresh:
        if cache_key in _MEMORY_CACHE:
            return _MEMORY_CACHE[cache_key]
        cached = _load_cache(source_cache, source_rules, source_projects)
        if cached is not None:
            _MEMORY_CACHE[cache_key] = cached
            return cached

    index = build_project_index(rules_path=source_rules, projects_dir=source_projects)
    _write_cache(source_cache, index)
    _MEMORY_CACHE[cache_key] = index
    return index


def _extract_emails(mail: dict[str, Any]) -> list[str]:
    text = " ".join(
        str(mail.get(key, ""))
        for key in ("sender_email", "sender_display", "from", "to", "cc")
    )
    found = re.findall(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", text, re.IGNORECASE)
    seen: set[str] = set()
    emails: list[str] = []
    for email in found:
        normalized = email.casefold()
        if normalized not in seen:
            seen.add(normalized)
            emails.append(normalized)
    return emails


def _domain_for_email(email: str) -> str:
    return email.rsplit("@", 1)[-1].casefold() if "@" in email else ""


def _domain_matches(domain: str, rule_domain: str) -> bool:
    return domain == rule_domain or domain.endswith("." + rule_domain)


def _rule_match(
    rule: dict[str, Any],
    *,
    source: str,
    matched: str,
    confidence: float,
) -> Optional[MailProjectMatch]:
    project = str(rule.get("project") or "").strip()
    if not project:
        return None
    return MailProjectMatch(
        project=project,
        area=str(rule.get("area") or "").strip(),
        source=source,
        confidence=float(rule.get("confidence") or confidence),
        matched=matched,
    )


def _mail_text(mail: dict[str, Any]) -> str:
    headers = str(mail.get("all_headers") or "")
    thread_topic = ""
    for line in headers.replace("\r\n", "\n").split("\n"):
        if line.casefold().startswith("thread-topic:"):
            thread_topic = line.split(":", 1)[1].strip()
            break
    return " ".join(part for part in (str(mail.get("subject") or ""), thread_topic) if part)


def _code_pattern(code: str) -> re.Pattern[str]:
    parts = [part for part in re.split(r"[^a-z0-9]+", code.casefold()) if part]
    if not parts:
        parts = [re.escape(code.casefold())]
    joined = r"[\s_-]+".join(re.escape(part) for part in parts)
    return re.compile(rf"(?<![a-z0-9]){joined}(?![a-z0-9])", re.IGNORECASE)


def _match_project_codes(mail: dict[str, Any], index: dict[str, Any]) -> Optional[MailProjectMatch]:
    text = _mail_text(mail)
    if not text:
        return None

    matches: list[dict[str, Any]] = []
    for code, entry in (index.get("codes") or {}).items():
        if _code_pattern(code).search(text):
            matches.append(entry)
    for code, entries in (index.get("conflicts") or {}).items():
        if _code_pattern(code).search(text):
            matches.extend(entries)

    if not matches:
        return None

    projects = sorted({str(match.get("project") or "") for match in matches if match.get("project")})
    if len(projects) > 1:
        return MailProjectMatch(
            project="",
            area="",
            source="ambiguous_project_code",
            confidence=0.0,
            matched=", ".join(sorted(str(match.get("code") or "") for match in matches)),
            candidates=tuple(projects),
        )

    match = matches[0]
    source = str(match.get("source") or "project_code")
    return MailProjectMatch(
        project=str(match["project"]),
        area=str(match.get("area") or ""),
        source=source,
        confidence=DEFAULT_CONFIDENCE.get(source, 0.85),
        matched=str(match.get("code") or ""),
    )


def _normalize_subject_text(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", value.casefold()).strip()


def _match_subject_rules(mail: dict[str, Any], index: dict[str, Any]) -> Optional[MailProjectMatch]:
    text = _normalize_subject_text(_mail_text(mail))
    if not text:
        return None

    matches: list[MailProjectMatch] = []
    for rule in index.get("subject_rules") or []:
        phrases = _listify(rule.get("contains"))
        if not phrases:
            continue
        if all(_normalize_subject_text(phrase) in text for phrase in phrases):
            match = _rule_match(
                rule,
                source="subject_rule",
                matched=str(rule.get("id") or ", ".join(phrases)),
                confidence=DEFAULT_CONFIDENCE["subject_rule"],
            )
            if match:
                matches.append(match)

    projects = sorted({match.project for match in matches})
    if len(projects) > 1:
        return MailProjectMatch(
            project="",
            area="",
            source="ambiguous_subject_rule",
            confidence=0.0,
            matched=", ".join(match.matched for match in matches),
            candidates=tuple(projects),
        )
    return matches[0] if matches else None


def resolve_mail_project(
    mail: dict[str, Any],
    *,
    index: Optional[dict[str, Any]] = None,
    rules_path: Optional[Path] = None,
    projects_dir: Optional[Path] = None,
    cache_path: Optional[Path] = None,
    force_refresh: bool = False,
) -> Optional[MailProjectMatch]:
    project_index = index or load_project_index(
        rules_path=rules_path,
        projects_dir=projects_dir,
        cache_path=cache_path,
        force_refresh=force_refresh,
    )

    emails = _extract_emails(mail)
    for email in emails:
        rule = (project_index.get("addresses") or {}).get(email)
        if rule:
            match = _rule_match(
                rule,
                source="address",
                matched=email,
                confidence=DEFAULT_CONFIDENCE["address"],
            )
            if match:
                return match

    domain_rules = project_index.get("domains") or {}
    sorted_domains = sorted(domain_rules, key=len, reverse=True)
    for email in emails:
        email_domain = _domain_for_email(email)
        for domain in sorted_domains:
            if _domain_matches(email_domain, domain):
                match = _rule_match(
                    domain_rules[domain],
                    source="domain",
                    matched=domain,
                    confidence=DEFAULT_CONFIDENCE["domain"],
                )
                if match:
                    return match

    return _match_project_codes(mail, project_index) or _match_subject_rules(mail, project_index)


def duplicate_project_codes(index: Optional[dict[str, Any]] = None) -> dict[str, list[dict[str, Any]]]:
    project_index = index or load_project_index()
    return project_index.get("conflicts") or {}


def suggested_mail_codes(project: dict[str, Any]) -> list[str]:
    if _listify(project.get("mail_codes")):
        return []

    slug = str(project.get("slug") or "").strip()
    title = str(project.get("title") or "").strip()
    candidates = [_project_alias(slug)]
    if title:
        cleaned_title = re.sub(r"[^a-z0-9]+", "-", title.casefold()).strip("-")
        cleaned_title = re.sub(r"^project-", "", cleaned_title)
        candidates.append(cleaned_title)

    result: list[str] = []
    seen: set[str] = set()
    for candidate in candidates:
        if len(candidate) < 2 or candidate in seen:
            continue
        seen.add(candidate)
        result.append(candidate)
    return result[:2]
