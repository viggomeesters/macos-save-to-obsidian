# macos-mail-to-obsidian

Self-contained mail-to-obsidian tooling for macOS (Apple Mail + Outlook).

- Save selected mail to the Obsidian vault as interaction notes.
- Keep conversation threading and basic deduplication.
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
cd /Users/viggomeesters/Dev/macos-mail-to-obsidian
python3 scripts/save_mail.py --client apple
python3 scripts/save_mail.py --client outlook
```

Raycast script command wrapper:

```bash
scripts/raycast/save-mail.sh
```

## Notes

- Default behavior and logic are intentionally duplicated from `life-os-core` for isolation.
- This repo expects your normal Life OS vault layout for targets configured in `shared/brain_lib.py`.
