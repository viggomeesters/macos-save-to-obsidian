# Changelog

## [Unreleased]

- Initial repository onboarding baseline.
- Renamed the repository scope from `macos-mail-to-obsidian` to
  `macos-save-to-obsidian` so Apple Mail, Outlook, Teams, and WhatsApp can live
  under one app-to-vault capture toolkit.
- Added New Outlook capture path for `save-mail`:
  capture selected Outlook message via clipboard copy and parse metadata/body locally,
  so legacy AppleScript dependency is avoided when New Outlook is enabled.
- Improved New Outlook capture by adding a second clipboard-copy attempt and clearer
  failure handling when System Events/Accessibility permission blocks copy actions.
- Hardened New Outlook clipboard capture:
  copy results are now verified with a nonce-based clipboard change check, and stale/foreign
  clipboard content is rejected before being treated as a selected-message payload.
- New Outlook mail-coercion now rejects non-mail clipboard payloads (for example Raycast
  command output) to avoid writing placeholder notes when copy capture fails silently.
- Fixed a regression where non-message clipboard text was still returned when copy attempts
  didn't produce a valid Outlook message (marker/empty clipboard), causing misleading
  payloads and script-run failures.
- Added optional fallback for New Outlook (`SAVE_MAIL_OUTLOOK_LEGACY_FALLBACK=1`) to avoid hard
  failures when clipboard capture does not return a message payload. This keeps New Outlook
  mode as default, but allows explicit fallback to legacy AppleScript metadata/body extraction.
- Default Outlook capture now prefers no-UI behavior: New Outlook clipboard capture avoids
  app activation/frontmost changes unless explicitly enabled via
  `SAVE_MAIL_OUTLOOK_UI_ACTIONS=1`. Legacy fallback now also requires that flag.
- Raycast mail commands are split per app (`Save Apple Mail`, `Save Outlook Mail`)
  so Apple Mail capture never opens Outlook through auto-detection.
- Outlook inbox mail can now archive after successful save via a guarded UI action;
  `--no-archive` remains available for manual runs.
- Removed default menu-bar-copy click fallback in New Outlook capture to stop mouse-move side effects.
  Only when `SAVE_MAIL_OUTLOOK_UI_ACTIONS=1` is set, it may execute menu-based copy actions.
- Avoid Raycast-style generic failure when no mail is selected by treating selection-related
  runtime errors as non-fatal in normal `save-mail` output flow (still logged as warnings),
  so the command stays visible instead of surfacing as “Failed to run a script”.
- Bugfix: prevent `save-mail` from activating the legacy Outlook AppleScript
  runtime while New Outlook is enabled, because that can open Outlook without the
  expected account/profile. New Outlook now fails fast with an explicit message
  until a compatible capture path exists.
- New Outlook copy parsing is now more tolerant to non-standard clipboard
  payloads (multiple pasteboard formats, fuzzy headers, and body-only payloads) so
  valid selection captures no longer reject on strict header checks before
  coercion.

## [0.1.0] - 2026-06-16

- Standalone macOS Mail + Outlook capture utilities
- Shared logic extraction from life-os-core for self-contained execution
- Added script wrappers for Raycast usage
