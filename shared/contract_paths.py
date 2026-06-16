"""Resolve Life OS contract files from the vault.

Declarative contracts live in the Obsidian vault under ``system/contracts``.
Code repositories may keep bootstrap fallbacks, but they are not the source of
truth.
"""

from __future__ import annotations

import json
import os
from pathlib import Path

CORE_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_VAULT_ROOT = (
    Path.home()
    / "Library"
    / "Mobile Documents"
    / "iCloud~md~obsidian"
    / "Documents"
    / "vault"
)

CONTRACT_ALIASES: dict[str, Path] = {
    "life-os-schema": Path("life-os-schema.yaml"),
    "life-os-schema-doc": Path("life-os-schema.md"),
    "automation-registry": Path("automation/automation-registry.yaml"),
    "automation-contract-schema": Path("automation/automation-contract-schema.yaml"),
    "multiplier-vision": Path("vision/multiplier-vision.yaml"),
    "human-vision": Path("vision/human-vision.yaml"),
    "scorecard": Path("vision/scorecard.yaml"),
    "design-principes": Path("principles/design-principes.yaml"),
    "security-policy": Path("security/security-policy.yaml"),
    "skill-schema": Path("skills/skill-schema.yaml"),
    "mail-project-rules": Path("mail/mail-project-rules.yaml"),
    "week-schedule": Path("schedule/week-schedule.yaml"),
    "shared-module": Path("modules/shared.module.yaml"),
    "sap-index": Path("sap-index.yaml"),
    "paths-schema": Path("schemas/paths.schema.json"),
    "pipeline-state-schema": Path("schemas/pipeline-state.schema.json"),
}

LEGACY_REPO_PATHS: dict[str, Path] = {
    "life-os-schema": CORE_ROOT / "schema" / "life-os-schema.yaml",
    "life-os-schema-doc": CORE_ROOT / "schema" / "life-os-schema.md",
    "automation-registry": CORE_ROOT / "schema" / "automation-registry.yaml",
    "automation-contract-schema": CORE_ROOT
    / "schema"
    / "automation-contract-schema.yaml",
    "multiplier-vision": CORE_ROOT / "schema" / "multiplier-vision.yaml",
    "human-vision": CORE_ROOT / "schema" / "human-vision.yaml",
    "design-principes": CORE_ROOT / "schema" / "design-principes.yaml",
    "security-policy": CORE_ROOT / "schema" / "security-policy.yaml",
    "skill-schema": CORE_ROOT / "schema" / "skill-schema.yaml",
    "mail-project-rules": CORE_ROOT / "schema" / "mail-project-rules.yaml",
    "week-schedule": CORE_ROOT / "schema" / "week-schedule.yaml",
    "shared-module": CORE_ROOT / "shared" / "module.yaml",
    "paths-schema": CORE_ROOT / "schema" / "paths.schema.json",
    "pipeline-state-schema": CORE_ROOT / "schema" / "pipeline-state.schema.json",
}


def _paths_file() -> Path | None:
    env = os.environ.get("LIFE_OS_PATHS")
    candidates = [
        Path(env).expanduser() if env else None,
        Path.home() / ".config" / "life-os" / "paths.json",
        CORE_ROOT / "context" / "paths.json",
        CORE_ROOT / "paths.json",
    ]
    for candidate in candidates:
        if candidate and candidate.exists():
            return candidate
    return None


def vault_root() -> Path:
    """Resolve the vault root without importing brain_lib."""
    env = os.environ.get("VAULT_ROOT") or os.environ.get("LIFE_OS_VAULT_ROOT")
    if env:
        return Path(env).expanduser().resolve()

    paths = _paths_file()
    if paths:
        try:
            data = json.loads(paths.read_text(encoding="utf-8"))
            root = data.get("vault_root")
            if root:
                return Path(root).expanduser().resolve()
        except Exception:
            pass

    return DEFAULT_VAULT_ROOT.expanduser().resolve()


def contracts_dir() -> Path:
    env = os.environ.get("LIFE_OS_CONTRACTS_DIR")
    if env:
        return Path(env).expanduser().resolve()
    return vault_root() / "system" / "contracts"


def contract_relpath(name_or_rel: str | Path) -> Path:
    raw = Path(name_or_rel)
    if raw.is_absolute():
        return raw
    return CONTRACT_ALIASES.get(str(name_or_rel), raw)


def contract_candidates(
    name_or_rel: str | Path, *, include_legacy: bool = True
) -> list[Path]:
    """Return lookup candidates in source-of-truth order."""
    rel = contract_relpath(name_or_rel)
    if rel.is_absolute():
        return [rel]

    candidates: list[Path] = []
    alias_key = str(name_or_rel)
    if alias_key == "automation-registry":
        override = os.environ.get("LIFE_OS_AUTOMATION_REGISTRY")
        if override:
            candidates.append(Path(override).expanduser())

    candidates.append(contracts_dir() / rel)

    if include_legacy:
        if rel.name.startswith("automation-"):
            candidates.append(vault_root() / "system" / "automation" / rel.name)
        legacy = LEGACY_REPO_PATHS.get(alias_key)
        if legacy:
            candidates.append(legacy)

    deduped: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        key = str(candidate)
        if key not in seen:
            deduped.append(candidate)
            seen.add(key)
    return deduped


def contract_path(
    name_or_rel: str | Path,
    *,
    must_exist: bool = False,
    include_legacy: bool = True,
) -> Path:
    """Resolve a contract path, preferring the vault.

    If ``must_exist`` is false and no candidate exists, this returns the vault
    target path so writers know where the contract should live.
    """
    candidates = contract_candidates(name_or_rel, include_legacy=include_legacy)
    for candidate in candidates:
        if candidate.exists():
            return candidate
    if must_exist:
        rendered = "\n".join(f"- {candidate}" for candidate in candidates)
        raise FileNotFoundError(
            f"Life OS contract not found: {name_or_rel}\n{rendered}"
        )
    return candidates[0]
