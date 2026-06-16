#!/usr/bin/env python3
"""Train mail project detection rules from project metadata."""

from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "shared"))

import brain_lib  # noqa: E402
from mail_project_rules import (  # noqa: E402
    duplicate_project_codes,
    load_project_index,
    suggested_mail_codes,
)

LOG_PATH = ROOT / "context" / "observability" / "save-mail-trainer.jsonl"


def _listify(value: Any) -> list[str]:
    if value is None or value == "":
        return []
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    return [item.strip() for item in str(value).split(",") if item.strip()]


def _write_log(event: dict[str, Any]) -> None:
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    payload = {**event, "ts": datetime.now().isoformat(timespec="seconds")}
    with LOG_PATH.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(payload, ensure_ascii=False, sort_keys=True) + "\n")


def _project_label(entry: dict[str, Any]) -> str:
    title = entry.get("title") or entry.get("project")
    return f"{entry.get('project')} ({title})"


def _apply_mail_codes(entry: dict[str, Any], codes: list[str]) -> bool:
    path_value = str(entry.get("path") or "")
    if not path_value:
        return False
    path = Path(path_value)
    fm, _ = brain_lib.read_note(path)
    if not fm:
        return False

    current = _listify(fm.get("mail_codes"))
    merged = list(current)
    for code in codes:
        if code not in merged:
            merged.append(code)

    if merged == current:
        return False
    if not brain_lib.update_frontmatter(path, {"mail_codes": merged}):
        return False

    _write_log(
        {
            "action": "add_mail_codes",
            "project": entry.get("project"),
            "path": str(path),
            "added": [code for code in merged if code not in current],
        }
    )
    return True


def print_conflicts(index: dict[str, Any]) -> int:
    conflicts = duplicate_project_codes(index)
    if not conflicts:
        print("Geen dubbele projectcodes gevonden.")
        return 0

    print("Dubbele projectcodes:\n")
    for code, entries in sorted(conflicts.items()):
        print(f"{code}:")
        for entry in entries:
            suggestions = ", ".join(suggested_mail_codes(entry)) or "-"
            print(f"  - {_project_label(entry)}")
            print(f"    voorstel mail_codes: {suggestions}")
        print()
    return len(conflicts)


def apply_conflict_suggestions(index: dict[str, Any], *, assume_yes: bool = False) -> int:
    conflicts = duplicate_project_codes(index)
    changed = 0
    for code, entries in sorted(conflicts.items()):
        print(f"\n{code}:")
        for entry in entries:
            suggestions = suggested_mail_codes(entry)
            if not suggestions:
                continue
            print(f"  {_project_label(entry)}")
            print(f"  voorstel: {', '.join(suggestions)}")
            should_apply = assume_yes
            if not assume_yes:
                answer = input("  Toevoegen aan mail_codes? [y/N] ").strip().lower()
                should_apply = answer in {"y", "yes", "j", "ja"}
            if should_apply and _apply_mail_codes(entry, suggestions):
                changed += 1
                print("  bijgewerkt")
            elif should_apply:
                print("  geen wijziging")
    if changed:
        load_project_index(force_refresh=True)
    return changed


def interactive_menu() -> int:
    while True:
        print("\nSave Mail Trainer")
        print("1. Dubbele projectcodes bekijken")
        print("2. Voorgestelde mail_codes toepassen")
        print("3. Projectindex-cache verversen")
        print("q. Stop")
        choice = input("> ").strip().lower()
        if choice == "1":
            print_conflicts(load_project_index())
        elif choice == "2":
            changed = apply_conflict_suggestions(load_project_index())
            print(f"\n{changed} projectnote(s) bijgewerkt.")
        elif choice == "3":
            load_project_index(force_refresh=True)
            print("Projectindex-cache ververst.")
        elif choice in {"q", "quit", "exit", ""}:
            return 0
        else:
            print("Onbekende keuze.")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--conflicts", action="store_true", help="Show duplicate project codes")
    parser.add_argument(
        "--apply-conflict-suggestions",
        action="store_true",
        help="Add suggested mail_codes for duplicate project codes",
    )
    parser.add_argument("--yes", action="store_true", help="Do not prompt when applying")
    parser.add_argument("--refresh-cache", action="store_true", help="Rebuild project index cache")
    parser.add_argument("--json", action="store_true", dest="json_output")
    args = parser.parse_args(argv)

    if args.refresh_cache:
        index = load_project_index(force_refresh=True)
        if args.json_output:
            print(json.dumps({"status": "ok", "codes": len(index.get("codes") or {})}))
        else:
            print("Projectindex-cache ververst.")
        return 0

    if args.conflicts:
        index = load_project_index()
        if args.json_output:
            print(json.dumps(duplicate_project_codes(index), ensure_ascii=False))
        else:
            print_conflicts(index)
        return 0

    if args.apply_conflict_suggestions:
        changed = apply_conflict_suggestions(load_project_index(), assume_yes=args.yes)
        print(f"{changed} projectnote(s) bijgewerkt.")
        return 0

    return interactive_menu()


if __name__ == "__main__":
    raise SystemExit(main())
