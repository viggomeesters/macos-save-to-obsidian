#!/usr/bin/env python3
"""Audit or backfill deterministic Save Mail contract fields.

This script is intentionally non-AI. It reports saved mail notes that are
missing the deterministic fields needed by a later Hermes AI mail enricher.
With --apply it backfills only fields that can be derived from existing
frontmatter.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[1]
REPO_SHARED = REPO_ROOT / "shared"
if str(REPO_SHARED) not in sys.path:
    sys.path.insert(0, str(REPO_SHARED))
if str(REPO_ROOT / "scripts") not in sys.path:
    sys.path.insert(0, str(REPO_ROOT / "scripts"))

import brain_lib
import vault_note_writer as vnw
from save_mail import (
    MAIL_CAPTURE_SOURCE,
    MAIL_CAPTURE_VERSION,
    MAIL_ENRICHMENT_STATUS_PENDING,
    MAIL_ENRICHMENT_VERSION,
    _email_domain,
    clean_mail_subject,
)

REQUIRED_FIELDS = (
    "capture_source",
    "capture_version",
    "enrichment_status",
    "enrichment_version",
    "raw_subject",
    "clean_subject",
    "sender_domain",
)


def split_note_content(content: str) -> tuple[dict[str, Any], str]:
    match = re.match(r"^---\n(.*?)\n---\n?(.*)$", content, flags=re.DOTALL)
    if not match:
        return {}, content
    fm = brain_lib.parse_frontmatter(content)
    return fm, match.group(2)


def iter_mail_notes(notes_dir: Path) -> list[Path]:
    if not notes_dir.exists():
        return []
    return sorted(notes_dir.rglob("*.md"))


def deterministic_backfill(frontmatter: dict[str, Any]) -> dict[str, Any]:
    patched: dict[str, Any] = {}
    raw_subject = str(
        frontmatter.get("raw_subject")
        or frontmatter.get("title")
        or frontmatter.get("subject")
        or "Geen onderwerp"
    )
    sender = str(frontmatter.get("from") or "")
    topics = frontmatter.get("topics") or []

    defaults = {
        "capture_source": MAIL_CAPTURE_SOURCE,
        "capture_version": MAIL_CAPTURE_VERSION,
        "enrichment_status": MAIL_ENRICHMENT_STATUS_PENDING,
        "enrichment_version": MAIL_ENRICHMENT_VERSION,
        "raw_subject": raw_subject,
        "clean_subject": clean_mail_subject(raw_subject),
        "sender_domain": _email_domain(sender),
        "topics_source": "legacy-backfill" if topics else "none",
        "topics_confidence": 0.5 if topics else 0.0,
    }
    for key, value in defaults.items():
        if key not in frontmatter or frontmatter.get(key) in (None, ""):
            patched[key] = value
    return patched


def audit_note(path: Path) -> dict[str, Any] | None:
    content = path.read_text(encoding="utf-8")
    frontmatter, body = split_note_content(content)
    if frontmatter.get("type") != "interaction" or frontmatter.get("category") != "mail":
        return None

    missing = [field for field in REQUIRED_FIELDS if field not in frontmatter]
    patch = deterministic_backfill(frontmatter)
    return {
        "path": str(path),
        "slug": frontmatter.get("slug") or path.stem,
        "missing": missing,
        "patch": patch,
        "body": body,
        "frontmatter": frontmatter,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--limit", type=int, default=50)
    parser.add_argument("--apply", action="store_true")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    findings: list[dict[str, Any]] = []
    for path in iter_mail_notes(brain_lib.cfg.vault_notes):
        finding = audit_note(path)
        if not finding:
            continue
        if finding["missing"] or finding["patch"]:
            findings.append(finding)
        if args.limit and len(findings) >= args.limit:
            break

    changed: list[str] = []
    if args.apply:
        for finding in findings:
            patch = finding["patch"]
            if not patch:
                continue
            frontmatter = dict(finding["frontmatter"])
            frontmatter.update(patch)
            vnw.write_vault_note(Path(finding["path"]), frontmatter, finding["body"])
            changed.append(finding["path"])

    summary = {
        "checked_dir": str(brain_lib.cfg.vault_notes),
        "findings": [
            {
                "path": item["path"],
                "slug": item["slug"],
                "missing": item["missing"],
                "patch_keys": sorted(item["patch"]),
            }
            for item in findings
        ],
        "changed": changed,
        "apply": args.apply,
    }

    if args.json:
        print(json.dumps(summary, ensure_ascii=False, indent=2))
    else:
        print(f"Mail notes needing deterministic contract fields: {len(findings)}")
        for item in summary["findings"]:
            print(f"- {item['slug']}: missing={item['missing']} patch={item['patch_keys']}")
        if args.apply:
            print(f"Changed: {len(changed)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
