# Contributing

Thanks for helping improve `macos-save-to-obsidian`.

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

## Development

- Make focused changes.
- Keep scripts self-contained and avoid new external dependencies unless needed.
- Keep platform-specific AppleScript changes in `scripts/*.py`.
- Run tests for touched behavior:

```bash
pytest
```

## Pull requests

- Describe what changed and why.
- Include manual verification notes (platform/app versions, command used).
- Keep commits small and scoped.
