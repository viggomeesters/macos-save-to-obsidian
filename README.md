# macos-save-to-obsidian

Self-contained macOS capture tooling for saving app content to Obsidian.

Current channels:

- Apple Mail -> Obsidian
- Outlook -> Obsidian

Planned channels:

- Microsoft Teams -> Obsidian
- WhatsApp -> Obsidian

Current mail capture behavior:

- Save selected mail to the Obsidian vault as interaction notes.
- Keep conversation threading and basic deduplication.
- Archive inbox mail after successful save when the channel supports it.
- Duplicate of the existing `life-os-core` save-mail stack, scoped to this repo.

This repo is code-self-contained: it contains the full mail capture logic and does not require `life-os-core` runtime imports.

## Usage

Install:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

```bash
python3 scripts/save_mail.py
python3 scripts/save_mail.py --client outlook
python3 scripts/save_mail.py --client auto
python3 scripts/save_mail.py --help
```

Run de runner direct:

```bash
cd /Users/viggomeesters/Dev/macos-save-to-obsidian
python3 scripts/save_mail.py --client apple
python3 scripts/save_mail.py --client outlook
```

Audit deterministic mail contract fields:

```bash
python3 scripts/audit_mail_contract.py --limit 25
python3 scripts/audit_mail_contract.py --limit 25 --json
python3 scripts/audit_mail_contract.py --limit 25 --apply
```

Raycast script command wrapper:

```bash
scripts/raycast/save-apple-mail.sh
scripts/raycast/save-outlook-mail.sh
```

## Mail enrichment handoff

`save_mail.py` owns fast, deterministic capture. It does not call AI. New mail
notes expose a stable handoff contract for a future Hermes AI mail enricher:

- Selection: `type: interaction`, `category: mail`, `capture_source: save-mail`,
  `enrichment_status: pending`.
- Stable capture fields: `capture_version`, `raw_subject`, `clean_subject`,
  `sender_domain`, `mail_link`, `mail_client`, `source_account`,
  `source_mailbox`.
- Deterministic routing provenance: `entity_source`, `entity_confidence`,
  `project_source`, `project_confidence`, `project_slug`, `topics_source`,
  `topics_confidence`.
- Thread grouping fields when available: `thread_id`, `thread_source`,
  `root_message_id`, `in_reply_to`, `references`, `thread_topic`,
  `thread_index_root`.

The Hermes AI mail enricher should write bounded enrichment output and then set
`enrichment_status: enriched` or `needs_review`, plus its own timestamp/version
metadata. It should not remove raw capture fields.

## Notes

- Default behavior and logic are intentionally duplicated from `life-os-core` for isolation.
- This repo expects your normal Life OS vault layout for targets configured in `shared/brain_lib.py`.
