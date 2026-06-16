# macos-mail-to-obsidian

Self-contained mail-to-obsidian tooling for macOS (Apple Mail + Outlook).

- Save selected mail to the Obsidian vault as interaction notes.
- Keep conversation threading and basic deduplication.
- Duplicate of the existing `life-os-core` save-mail stack, scoped to this repo.

## Usage

```bash
python3 scripts/save_mail.py
python3 scripts/save_mail.py --client outlook
python3 scripts/save_mail.py --client auto
python3 scripts/save_mail.py --help
```

Raycast script command wrapper:

```bash
scripts/raycast/save-mail.sh
```

## Notes

- Default behavior and logic are intentionally duplicated from `life-os-core` for isolation.
- This repo expects your normal Life OS vault layout for targets configured in `shared/brain_lib.py`.
