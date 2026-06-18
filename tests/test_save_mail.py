"""Tests for save_mail.py calendar invite detection."""

from __future__ import annotations

import sys
import sqlite3
from types import SimpleNamespace
from contextlib import ExitStack
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "shared"))
sys.path.insert(0, str(REPO_ROOT / "scripts"))

import mail_applescript
import mail_outlook_applescript
import mail_project_rules
import save_mail
from save_mail import (
    conversation_fetch_hints,
    conversation_lookup_ids,
    conversation_message_ids,
    normalize_subject,
    parse_mail_headers,
    render_thread_timeline,
    replace_thread_timeline_block,
    save_selected_mail,
    select_conversation_mails,
    should_create_follow_up_task,
    is_calendar_invite,
    update_conversation_thread_timeline,
    update_note_thread_timeline,
)


def _mail(
    message_id: str,
    subject: str,
    *,
    sender: str = "person@example.com",
    to: str = "viggomeesters@icloud.com",
    cc: str = "",
    headers: str = "",
    date: str = "2026-05-01 10:00:00",
    mailbox_type: str = "inbox",
    mailbox: str = "Inbox",
) -> dict:
    return {
        "subject": subject,
        "sender_email": sender,
        "sender_display": sender,
        "date_str": date,
        "message_id": message_id,
        "account": "iCloud",
        "mailbox": mailbox,
        "to": to,
        "cc": cc,
        "att_count": "0",
        "att_names": "",
        "is_flagged": False,
        "all_headers": headers,
        "mailbox_type": mailbox_type,
        "body": "",
    }


class TestConversationDetection:
    def test_normalize_subject_strips_reply_prefixes(self) -> None:
        assert normalize_subject("Re: AW: Antw: Project Update") == "project update"
        assert normalize_subject("FW: SV: ITSM-1601") == "itsm-1601"

    def test_parse_headers_unfolds_values(self) -> None:
        headers = parse_mail_headers(
            "Message-ID: <a@example.com>\n"
            "References: <root@example.com>\n"
            " <parent@example.com>\n"
            "Thread-Topic: Project Update\n"
        )

        assert headers["message-id"] == "<a@example.com>"
        assert headers["references"] == "<root@example.com> <parent@example.com>"
        assert headers["thread-topic"] == "Project Update"

    def test_conversation_message_ids_from_reply_headers(self) -> None:
        mail = _mail(
            "c@example.com",
            "Re: Project Update",
            headers=(
                "Message-ID: <c@example.com>\n"
                "In-Reply-To: <b@example.com>\n"
                "References: <a@example.com> <b@example.com>\n"
            ),
        )

        assert conversation_message_ids(mail) == [
            "c@example.com",
            "b@example.com",
            "a@example.com",
        ]

    def test_references_with_multiple_found_mails_is_conversation(self) -> None:
        root = _mail(
            "a@example.com",
            "Project Update",
            headers="Message-ID: <a@example.com>\n",
            date="2026-05-01 09:00:00",
        )
        reply = _mail(
            "b@example.com",
            "Re: Project Update",
            headers=(
                "Message-ID: <b@example.com>\n"
                "In-Reply-To: <a@example.com>\n"
                "References: <a@example.com>\n"
            ),
            date="2026-05-01 10:00:00",
        )

        selection = select_conversation_mails(reply, [root, reply])

        assert selection["is_conversation"]
        assert selection["reason"] == "message-id"
        assert [mail["message_id"] for mail in selection["mails"]] == [
            "a@example.com",
            "b@example.com",
        ]

    def test_newsletter_without_reply_headers_stays_single(self) -> None:
        selected = _mail(
            "one@example.com",
            "Latest News from SAP",
            sender="donotreply@sappartnerupdate.com",
            to="viggo.meesters@mccoy-partners.com",
        )
        other = _mail(
            "two@example.com",
            "Latest News from SAP",
            sender="donotreply@sappartnerupdate.com",
            to="viggo.meesters@mccoy-partners.com",
        )

        selection = select_conversation_mails(selected, [selected, other])

        assert not selection["is_conversation"]

    def test_same_reply_subject_different_participants_stays_single(self) -> None:
        selected = _mail(
            "one@example.com",
            "Re: Budget",
            sender="anne@example.com",
            to="viggomeesters@icloud.com",
        )
        other = _mail(
            "two@example.com",
            "Re: Budget",
            sender="other@example.com",
            to="viggomeesters@icloud.com",
        )

        selection = select_conversation_mails(selected, [selected, other])

        assert not selection["is_conversation"]

    def test_sent_and_received_mail_in_same_thread_is_conversation(self) -> None:
        sent = _mail(
            "a@example.com",
            "Project Update",
            sender="viggomeesters@icloud.com",
            to="person@example.com",
            headers="Message-ID: <a@example.com>\n",
            mailbox_type="sent",
            mailbox="Sent Messages",
            date="2026-05-01 09:00:00",
        )
        received = _mail(
            "b@example.com",
            "Re: Project Update",
            headers=(
                "Message-ID: <b@example.com>\n"
                "In-Reply-To: <a@example.com>\n"
                "References: <a@example.com>\n"
            ),
            date="2026-05-01 10:00:00",
        )

        selection = select_conversation_mails(received, [sent, received])

        assert selection["is_conversation"]
        assert [mail["mailbox_type"] for mail in selection["mails"]] == [
            "sent",
            "inbox",
        ]

    def test_fetch_hints_enable_subject_for_reply_headers(self) -> None:
        selected = _mail(
            "b@example.com",
            "Re: Project Update",
            headers=(
                "Message-ID: <b@example.com>\n"
                "In-Reply-To: <a@example.com>\n"
            ),
        )

        hints = conversation_fetch_hints(selected)

        assert hints["include_subject"]
        assert hints["subject_hint"] == "Project Update"
        assert hints["message_ids"] == ["b@example.com", "a@example.com"]

    def test_fetch_hints_do_not_enable_subject_for_forward_prefix(self) -> None:
        selected = _mail(
            "b@example.com",
            "FW: OZI zomerbbq",
            headers="Message-ID: <b@example.com>\n",
        )

        hints = conversation_fetch_hints(selected)

        assert not hints["include_subject"]
        assert hints["subject_hint"] == "OZI zomerbbq"
        assert hints["message_ids"] == ["b@example.com"]

    def test_outlook_thread_headers_alone_do_not_enable_subject_search(self) -> None:
        selected = _mail(
            "a@example.com",
            "Project Update",
            headers=(
                "Message-ID: <a@example.com>\n"
                "Thread-Topic: Project Update\n"
                "Thread-Index: Acabcd1234567890abcd1234567890abcd1234567890\n"
            ),
        )

        hints = conversation_fetch_hints(selected)

        assert not hints["include_subject"]
        assert hints["message_ids"] == ["a@example.com"]

    def test_lookup_ids_excludes_selected_message(self) -> None:
        selected = _mail(
            "b@example.com",
            "Re: Project Update",
            headers=(
                "Message-ID: <b@example.com>\n"
                "In-Reply-To: <a@example.com>\n"
                "References: <a@example.com>\n"
            ),
        )

        assert conversation_lookup_ids(selected) == ["a@example.com"]

    def test_non_reply_references_do_not_trigger_lookup(self) -> None:
        selected = _mail(
            "b@example.com",
            "Newsletter",
            headers=(
                "Message-ID: <b@example.com>\n"
                "References: <a@example.com>\n"
            ),
        )

        hints = conversation_fetch_hints(selected)

        assert hints["message_ids"] == ["b@example.com", "a@example.com"]
        assert not hints["include_subject"]
        assert conversation_lookup_ids(selected) == []


class TestDeduplication:
    def test_empty_message_id_is_never_duplicate(
        self, tmp_path: Path, monkeypatch
    ) -> None:
        cache = tmp_path / "dedup.json"
        cache.write_text('{"": "bad-slug"}', encoding="utf-8")
        monkeypatch.setattr(save_mail, "_DEDUP_CACHE", cache)

        assert save_mail.is_duplicate("") is None

    def test_sqlite_lookup_escapes_like_wildcards(
        self, tmp_path: Path, monkeypatch
    ) -> None:
        db_path = tmp_path / ".brain-vault-index.sqlite"
        conn = sqlite3.connect(db_path)
        conn.execute(
            "CREATE TABLE notes (path TEXT, type TEXT, category TEXT, frontmatter_json TEXT)"
        )
        conn.execute(
            "INSERT INTO notes VALUES (?, ?, ?, ?)",
            (
                "10_notes/wildcard.md",
                "interaction",
                "mail",
                '{"mail_link":"message://<abcXexample@example.com>"}',
            ),
        )
        conn.execute(
            "INSERT INTO notes VALUES (?, ?, ?, ?)",
            (
                "10_notes/exact.md",
                "interaction",
                "mail",
                '{"mail_link":"message://<abc_example@example.com>"}',
            ),
        )
        conn.commit()
        conn.close()
        monkeypatch.setattr(save_mail, "VAULT_ROOT", tmp_path)
        monkeypatch.setattr(save_mail, "_DEDUP_CACHE", tmp_path / "missing.json")

        assert save_mail.is_duplicate("abc_example@example.com") == "exact"


class TestMailAppleScriptParsing:
    def test_parse_mail_record_recovers_message_id_from_headers(self) -> None:
        fields = [
            "inbox",
            "Subject",
            "person@example.com",
            "Person <person@example.com>",
            "2026-05-01 10:00:00",
            "",
            "iCloud",
            "Inbox",
            "me@example.com",
            "",
            "0",
            "",
            "false",
            "false",
            "Message-ID: <real@example.com>\n",
        ]

        parsed = mail_applescript._parse_mail_record("|||".join(fields))

        assert parsed is not None
        assert parsed["message_id"] == "real@example.com"


class TestThreadTimeline:
    def test_render_thread_timeline_marks_current(self) -> None:
        lines = render_thread_timeline(
            [
                {"slug": "20260504-1019-mail-reply", "direction": "sent"},
                {"slug": "20260501-1640-mail-root", "direction": "received"},
                {"slug": "20260504-1024-mail-followup", "direction": "received"},
            ],
            "20260504-1019-mail-reply",
        )

        assert lines == [
            "🔗 Thread:",
            "- Received [[20260501-1640-mail-root]]",
            "- Sent [[20260504-1019-mail-reply]] (current)",
            "- Received [[20260504-1024-mail-followup]]",
        ]

    def test_replace_old_inline_thread_with_timeline(self) -> None:
        body = "\n".join(
            [
                "# Subject",
                "[📩 Open in Mail](message://<id>)",
                "🔗 Thread: [[old-a]] → [[old-b]]",
                "",
                "---",
                "",
                "Mail body",
            ]
        )

        result = replace_thread_timeline_block(
            body,
            [
                {"slug": "20260501-1640-mail-root", "direction": "received"},
                {"slug": "20260504-1019-mail-reply", "direction": "sent"},
            ],
            current_slug="20260504-1019-mail-reply",
        )

        assert "→" not in result
        assert "🔗 Thread:\n- Received [[20260501-1640-mail-root]]" in result
        assert "- Sent [[20260504-1019-mail-reply]] (current)" in result
        assert "Mail body" in result

    def test_replace_multiline_thread_is_idempotent(self) -> None:
        entries = [
            {"slug": "20260501-1640-mail-root", "direction": "received"},
            {"slug": "20260504-1019-mail-reply", "direction": "sent"},
        ]
        body = "\n".join(
            [
                "# Subject",
                "[📩 Open in Mail](message://<id>)",
                "🔗 Thread:",
                "- Received [[stale-root]]",
                "- Sent [[stale-reply]] (current)",
                "",
                "---",
                "",
                "Mail body",
            ]
        )

        once = replace_thread_timeline_block(
            body,
            entries,
            current_slug="20260504-1019-mail-reply",
        )
        twice = replace_thread_timeline_block(
            once,
            entries,
            current_slug="20260504-1019-mail-reply",
        )

        assert once == twice
        assert once.count("🔗 Thread:") == 1
        assert "stale-root" not in once

    def test_update_note_thread_timeline_updates_frontmatter_and_body(
        self, tmp_path: Path
    ) -> None:
        path = tmp_path / "20260504-1019-mail-reply.md"
        path.write_text(
            "\n".join(
                [
                    "---",
                    "type: interaction",
                    "category: mail",
                    "created: 2026-05-04",
                    "slug: 20260504-1019-mail-reply",
                    "timestamp: 20260504-1019",
                    "area: work",
                    "direction: sent",
                    "thread: [old-thread]",
                    "---",
                    "",
                    "# Subject",
                    "[📩 Open in Mail](message://<id>)",
                    "📋 Task: [[task-slug]]",
                    "🔗 Thread: [[old-thread]]",
                    "",
                    "---",
                    "",
                    "Mail body",
                ]
            ),
            encoding="utf-8",
        )
        entries = [
            {"slug": "20260501-1640-mail-root", "direction": "received"},
            {"slug": "20260504-1019-mail-reply", "direction": "sent"},
            {"slug": "20260504-1024-mail-followup", "direction": "received"},
        ]

        assert update_note_thread_timeline(
            path,
            entries,
            current_slug="20260504-1019-mail-reply",
        )

        content = path.read_text(encoding="utf-8")
        assert (
            'thread: ["20260501-1640-mail-root", "20260504-1024-mail-followup"]'
            in content
        )
        assert "📋 Task: [[task-slug]]\n🔗 Thread:" in content
        assert "- Sent [[20260504-1019-mail-reply]] (current)" in content
        assert "Mail body" in content

    def test_update_conversation_thread_timeline_updates_all_notes(
        self, tmp_path: Path
    ) -> None:
        entries = [
            {
                "slug": "20260501-1640-mail-root",
                "direction": "received",
                "path": str(tmp_path / "20260501-1640-mail-root.md"),
            },
            {
                "slug": "20260504-1019-mail-reply",
                "direction": "sent",
                "path": str(tmp_path / "20260504-1019-mail-reply.md"),
            },
            {
                "slug": "20260504-1024-mail-followup",
                "direction": "received",
                "path": str(tmp_path / "20260504-1024-mail-followup.md"),
            },
        ]
        for entry in entries:
            slug = entry["slug"]
            Path(entry["path"]).write_text(
                "\n".join(
                    [
                        "---",
                        "type: interaction",
                        "category: mail",
                        "created: 2026-05-04",
                        f"slug: {slug}",
                        f"timestamp: {slug[:13]}",
                        "area: work",
                        f"direction: {entry['direction']}",
                        "---",
                        "",
                        f"# {slug}",
                        "[📩 Open in Mail](message://<id>)",
                        "",
                        "---",
                        "",
                        "Mail body",
                    ]
                ),
                encoding="utf-8",
            )

        assert update_conversation_thread_timeline(entries) == 3

        root = Path(entries[0]["path"]).read_text(encoding="utf-8")
        reply = Path(entries[1]["path"]).read_text(encoding="utf-8")
        followup = Path(entries[2]["path"]).read_text(encoding="utf-8")

        assert "- Received [[20260501-1640-mail-root]] (current)" in root
        assert "- Sent [[20260504-1019-mail-reply]] (current)" in reply
        assert "- Received [[20260504-1024-mail-followup]] (current)" in followup
        assert 'thread: ["20260504-1019-mail-reply", "20260504-1024-mail-followup"]' in root
        assert 'thread: ["20260501-1640-mail-root", "20260504-1024-mail-followup"]' in reply


class TestOutlookClipboardParsing:
    def test_rtf_clipboard_payload_coerces_to_outlook_record(self) -> None:
        raw = (
            r"{\rtf1\ansi From: Jane Example <jane@example.com>\par "
            r"Subject: Planning update\par "
            r"Date: 16 June 2026 12:00:00\par "
            r"To: Viggo <viggo@example.com>\par\par Body line one\par Body line two}"
        )

        record = mail_outlook_applescript._coerce_outlook_record_from_clipboard(raw)

        assert record is not None
        assert record["subject"] == "Planning update"
        assert record["sender_email"] == "jane@example.com"
        assert record["to"] == "viggo@example.com"
        assert "Body line one" in record["body"]

    def test_tab_separated_clipboard_headers_are_parsed(self) -> None:
        raw = "\n".join(
            [
                "From\tJane Example <jane@example.com>",
                "Sent\t16 June 2026 12:00:00",
                "To\tViggo <viggo@example.com>",
                "Subject\tPlanning update",
                "",
                "Body line",
            ]
        )

        record = mail_outlook_applescript._coerce_outlook_record_from_clipboard(raw)

        assert record is not None
        assert record["subject"] == "Planning update"
        assert record["sender_email"] == "jane@example.com"
        assert record["date_str"] == "16 June 2026 12:00:00"
        assert record["body"] == "Body line"

    def test_clipboard_reader_uses_rtf_candidate_when_plain_text_is_not_mail(
        self,
    ) -> None:
        raw_rtf = (
            r"{\rtf1\ansi From: Jane Example <jane@example.com>\par "
            r"Subject: Planning update}"
        )

        def fake_run(cmd: list[str], **kwargs) -> SimpleNamespace:
            if cmd == ["/usr/bin/pbpaste", "-Prefer", "rtf"]:
                return SimpleNamespace(returncode=0, stdout=raw_rtf)
            return SimpleNamespace(returncode=0, stdout="not a mail payload")

        with patch("mail_outlook_applescript.subprocess.run", side_effect=fake_run):
            text = mail_outlook_applescript._clipboard_paste_prefer_rich()

        assert "From: Jane Example <jane@example.com>" in text
        assert "Subject: Planning update" in text

    def test_new_outlook_accessibility_row_parses_message_metadata(self) -> None:
        row = (
            "Jane Example, Planning update,     16/06/2026,        "
            "Hi Viggo, this is the message preview."
        )

        record = mail_outlook_applescript._parse_new_outlook_ax_row(
            row,
            window_title="Inbox • jane@example.com",
        )

        assert record is not None
        assert record["subject"] == "Planning update"
        assert record["sender_display"] == "Jane Example"
        assert record["date_str"] == "16/06/2026"
        assert record["mailbox"] == "Inbox"
        assert record["account"] == "jane@example.com"
        assert record["body"] == "Hi Viggo, this is the message preview."
        assert record["capture_source"] == "outlook-accessibility"

    def test_new_outlook_accessibility_row_parses_conversation_prefix(self) -> None:
        row = (
            "7 messages, Leo van Horrik, Alba Colitti, Gebruiker Viggo Meesters, "
            "Projectvraag,     10/06/2026,        Hi Alba, kun jij kijken?"
        )

        record = mail_outlook_applescript._parse_new_outlook_ax_row(row)

        assert record is not None
        assert record["sender_display"] == "Leo van Horrik, Alba Colitti, Gebruiker Viggo Meesters"
        assert record["subject"] == "Projectvraag"

    def test_raycast_output_is_not_treated_as_outlook_mail(self) -> None:
        raw = "Running: scripts/raycast/save-outlook-mail.sh\n▶ Script: scripts/save_mail.py"

        assert not mail_outlook_applescript._looks_like_message_copy(raw)
        assert mail_outlook_applescript._coerce_outlook_record_from_clipboard(raw) is None


class TestMailProjectDetection:
    def test_short_project_code_does_not_match_common_word_in_subject(self) -> None:
        index = {
            "addresses": {},
            "domains": {},
            "subject_rules": [],
            "codes": {
                "je": {
                    "project": "2025-12-journal-extension",
                    "area": "work",
                    "code": "je",
                    "source": "project_code",
                }
            },
            "conflicts": {},
        }

        match = mail_project_rules.resolve_mail_project(
            {
                "subject": "AskIT: Je Request REQ0116499 is Afgehandeld",
                "sender_email": "askit servicenow",
                "sender_display": "askit servicenow",
            },
            index=index,
        )

        assert match is None

    def test_short_explicit_mail_code_can_still_match_subject(self) -> None:
        index = {
            "addresses": {},
            "domains": {},
            "subject_rules": [],
            "codes": {
                "je": {
                    "project": "2025-12-journal-extension",
                    "area": "work",
                    "code": "je",
                    "source": "mail_code",
                }
            },
            "conflicts": {},
        }

        match = mail_project_rules.resolve_mail_project(
            {
                "subject": "Project update JE",
                "sender_email": "person@example.com",
                "sender_display": "Person",
            },
            index=index,
        )

        assert match is not None
        assert match.project == "2025-12-journal-extension"
        assert match.source == "mail_code"


class TestConversationSaveFlow:
    def test_explicit_outlook_client_routes_through_outlook_adapter(self) -> None:
        selected = _mail("outlook-id", "Outlook subject")
        selected.pop("mail_client", None)
        adapter_calls: list[str] = []
        created: list[dict] = []

        def fake_archive(*args, **kwargs) -> str:
            adapter_calls.append("archive")
            return "OK"

        fake_adapter = SimpleNamespace(
            get_selected_mail_headers=lambda: [selected],
            fetch_mail_body=lambda *args, **kwargs: "outlook body",
            save_attachments=lambda *args, **kwargs: [],
            archive_mail=fake_archive,
            fetch_all_mailboxes=lambda since_days=7: {
                "inbox": [],
                "sent": [],
                "deleted": [],
            },
        )

        def fake_load_mail_client(client: str):
            assert client == "outlook"
            return fake_adapter

        def fake_create_note(mail: dict, **kwargs) -> dict:
            created.append(dict(mail))
            return {"slug": "slug-outlook", "thread": []}

        with ExitStack() as stack:
            stack.enter_context(
                patch("save_mail._load_mail_client", side_effect=fake_load_mail_client)
            )
            stack.enter_context(patch("save_mail.is_duplicate", return_value=None))
            stack.enter_context(
                patch("save_mail.create_mail_note", side_effect=fake_create_note)
            )
            stack.enter_context(patch("save_mail._write_log"))

            outcome = save_selected_mail(client="outlook", verbose=False)

        assert outcome["mode"] == "single"
        assert outcome["results"][0]["slug"] == "slug-outlook"
        assert created[0]["mail_client"] == "outlook"
        assert created[0]["body"] == "outlook body"
        assert outcome["results"][0]["archived"] == "OK"
        assert adapter_calls == ["archive"]

    def test_explicit_outlook_client_respects_no_archive(self) -> None:
        selected = _mail("outlook-id", "Outlook subject")
        selected.pop("mail_client", None)
        adapter_calls: list[str] = []

        fake_adapter = SimpleNamespace(
            get_selected_mail_headers=lambda: [selected],
            fetch_mail_body=lambda *args, **kwargs: "outlook body",
            save_attachments=lambda *args, **kwargs: [],
            archive_mail=lambda *args, **kwargs: adapter_calls.append("archive"),
            fetch_all_mailboxes=lambda since_days=7: {
                "inbox": [],
                "sent": [],
                "deleted": [],
            },
        )

        def fake_load_mail_client(client: str):
            assert client == "outlook"
            return fake_adapter

        with ExitStack() as stack:
            stack.enter_context(
                patch("save_mail._load_mail_client", side_effect=fake_load_mail_client)
            )
            stack.enter_context(patch("save_mail.is_duplicate", return_value=None))
            stack.enter_context(
                patch("save_mail.create_mail_note", return_value={"slug": "slug-outlook"})
            )
            stack.enter_context(patch("save_mail._write_log"))

            outcome = save_selected_mail(
                client="outlook",
                no_archive=True,
                verbose=False,
            )

        assert outcome["results"][0]["slug"] == "slug-outlook"
        assert "archived" not in outcome["results"][0]
        assert adapter_calls == []

    def test_auto_client_prefers_outlook_when_outlook_has_selection(self) -> None:
        outlook_mail = _mail("outlook-id", "Outlook subject")
        outlook_adapter = SimpleNamespace(
            is_outlook_running=lambda: True,
            get_selected_mail_headers=lambda: [outlook_mail],
        )
        apple_adapter = SimpleNamespace(
            is_mail_running=lambda: True,
            get_selected_mail_headers=lambda: [_mail("apple-id", "Apple subject")],
        )

        def fake_load_mail_client(client: str):
            return {"outlook": outlook_adapter, "apple": apple_adapter}[client]

        with patch("save_mail._load_mail_client", side_effect=fake_load_mail_client):
            client, mails = save_mail.get_selected_headers_for_client("auto")

        assert client == "outlook"
        assert mails == [outlook_mail]

    def test_auto_client_reports_clear_error_when_no_selection_exists(self) -> None:
        outlook_adapter = SimpleNamespace(
            is_outlook_running=lambda: True,
            get_selected_mail_headers=lambda: (_ for _ in ()).throw(
                RuntimeError("No selected Outlook mail")
            ),
        )
        apple_adapter = SimpleNamespace(
            is_mail_running=lambda: True,
            get_selected_mail_headers=lambda: (_ for _ in ()).throw(
                RuntimeError("No selected mail")
            ),
        )

        def fake_load_mail_client(client: str):
            return {"outlook": outlook_adapter, "apple": apple_adapter}[client]

        with patch("save_mail._load_mail_client", side_effect=fake_load_mail_client):
            try:
                save_mail.get_selected_headers_for_client("auto")
            except RuntimeError as exc:
                message = str(exc)
            else:
                raise AssertionError("Expected RuntimeError")

        assert "No selected mail found in Outlook or Apple Mail" in message
        assert "outlook: No selected Outlook mail" in message
        assert "apple: No selected mail" in message

    def test_auto_conversation_saves_chronologically_and_archives_inbox_only(
        self,
    ) -> None:
        sent = _mail(
            "a@example.com",
            "Project Update",
            sender="viggomeesters@icloud.com",
            to="person@example.com",
            headers="Message-ID: <a@example.com>\n",
            mailbox_type="sent",
            mailbox="Sent Messages",
            date="2026-05-01 09:00:00",
        )
        selected = _mail(
            "b@example.com",
            "Re: Project Update",
            headers=(
                "Message-ID: <b@example.com>\n"
                "In-Reply-To: <a@example.com>\n"
                "References: <a@example.com>\n"
            ),
            date="2026-05-01 10:00:00",
        )
        created: list[tuple[str, list[str]]] = []

        def fake_create_note(
            mail: dict,
            create_task: bool = False,
            extra_thread_slugs: list[str] | None = None,
        ) -> dict:
            slug = f"slug-{mail['message_id'][0]}"
            created.append((mail["message_id"], list(extra_thread_slugs or [])))
            return {"slug": slug, "thread": list(extra_thread_slugs or [])}

        with ExitStack() as stack:
            stack.enter_context(
                patch("save_mail.get_selected_mail_headers", return_value=[selected])
            )
            stack.enter_context(
                patch(
                    "save_mail.fetch_conversation_candidates",
                    return_value=[sent, selected],
                )
            )
            stack.enter_context(patch("save_mail.is_duplicate", return_value=None))
            stack.enter_context(patch("save_mail.fetch_mail_body", return_value="body"))
            stack.enter_context(
                patch("save_mail.create_mail_note", side_effect=fake_create_note)
            )
            archive_mail = stack.enter_context(
                patch("save_mail.archive_mail", return_value="OK")
            )
            stack.enter_context(patch("save_mail._write_log"))
            outcome = save_selected_mail(verbose=False)

        assert outcome["mode"] == "conversation"
        assert [result["slug"] for result in outcome["results"]] == [
            "slug-a",
            "slug-b",
        ]
        assert created == [
            ("a@example.com", []),
            ("b@example.com", ["slug-a"]),
        ]
        archive_mail.assert_called_once_with(
            "b@example.com",
            "iCloud",
            mailbox="Inbox",
        )

    def test_single_mail_without_thread_hints_skips_conversation_fetch(self) -> None:
        selected = _mail("one@example.com", "Plain subject")

        with ExitStack() as stack:
            stack.enter_context(
                patch("save_mail.get_selected_mail_headers", return_value=[selected])
            )
            fetch_candidates = stack.enter_context(
                patch("save_mail.fetch_conversation_candidates")
            )
            process_mail = stack.enter_context(
                patch(
                    "save_mail.process_mail_record",
                    return_value={"status": "saved", "slug": "slug-one"},
                )
            )

            outcome = save_selected_mail(verbose=False)

        assert outcome["mode"] == "single"
        assert outcome["reason"] == "single"
        fetch_candidates.assert_not_called()
        process_mail.assert_called_once()

    def test_forwarded_mail_without_thread_hints_skips_conversation_fetch(self) -> None:
        selected = _mail(
            "one@example.com",
            "FW: OZI zomerbbq",
            headers="Message-ID: <one@example.com>\n",
        )

        with ExitStack() as stack:
            stack.enter_context(
                patch("save_mail.get_selected_mail_headers", return_value=[selected])
            )
            fetch_candidates = stack.enter_context(
                patch("save_mail.fetch_conversation_candidates")
            )
            process_mail = stack.enter_context(
                patch(
                    "save_mail.process_mail_record",
                    return_value={"status": "saved", "slug": "slug-one"},
                )
            )

            outcome = save_selected_mail(verbose=False)

        assert outcome["mode"] == "single"
        assert outcome["reason"] == "single"
        fetch_candidates.assert_not_called()
        process_mail.assert_called_once()

    def test_non_reply_references_skip_conversation_fetch(self) -> None:
        selected = _mail(
            "b@example.com",
            "Gefeliciteerd met Groningen Go Live",
            headers=(
                "Message-ID: <b@example.com>\n"
                "References: <a@example.com>\n"
            ),
        )

        with ExitStack() as stack:
            stack.enter_context(
                patch("save_mail.get_selected_mail_headers", return_value=[selected])
            )
            fetch_candidates = stack.enter_context(
                patch("save_mail.fetch_conversation_candidates")
            )
            process_mail = stack.enter_context(
                patch(
                    "save_mail.process_mail_record",
                    return_value={"status": "saved", "slug": "slug-one"},
                )
            )

            outcome = save_selected_mail(verbose=False)

        assert outcome["mode"] == "single"
        assert outcome["reason"] == "single"
        fetch_candidates.assert_not_called()
        process_mail.assert_called_once()

    def test_multi_selected_mail_without_thread_hints_saves_selection(
        self,
    ) -> None:
        sent = _mail(
            "sent@example.com",
            "Sent follow-up",
            sender="viggomeesters@icloud.com",
            to="person@example.com",
            mailbox_type="sent",
            mailbox="Sent Messages",
        )
        deleted_calendar = _mail(
            "invite@example.com",
            "Accepted: Planning",
            sender="person@example.com",
            to="viggomeesters@icloud.com",
            mailbox_type="deleted",
            mailbox="Deleted Messages",
        )

        with ExitStack() as stack:
            stack.enter_context(
                patch(
                    "save_mail.get_selected_mail_headers",
                    return_value=[sent, deleted_calendar],
                )
            )
            fetch_candidates = stack.enter_context(
                patch("save_mail.fetch_conversation_candidates")
            )
            process_mail = stack.enter_context(
                patch(
                    "save_mail.process_mail_record",
                    side_effect=[
                        {"status": "saved", "slug": "slug-sent"},
                        {"status": "saved", "slug": "slug-calendar"},
                    ],
                )
            )

            outcome = save_selected_mail(verbose=False)

        assert outcome["mode"] == "selection"
        assert outcome["reason"] == "selected-conversation"
        assert [r["slug"] for r in outcome["results"]] == [
            "slug-sent",
            "slug-calendar",
        ]
        fetch_candidates.assert_not_called()
        assert process_mail.call_count == 2
        assert process_mail.call_args_list[0].args[0] is sent
        assert process_mail.call_args_list[1].args[0] is deleted_calendar
        assert process_mail.call_args_list[0].kwargs["single_mode"] is True
        assert process_mail.call_args_list[1].kwargs["single_mode"] is True

    def test_multi_selected_mail_expands_conversation_when_thread_hints_exist(
        self,
    ) -> None:
        root = _mail(
            "root@example.com",
            "Project Update",
            headers="Message-ID: <root@example.com>\n",
            date="2026-05-01 09:00:00",
        )
        selected_reply = _mail(
            "reply@example.com",
            "Re: Project Update",
            headers=(
                "Message-ID: <reply@example.com>\n"
                "In-Reply-To: <root@example.com>\n"
                "References: <root@example.com>\n"
            ),
            date="2026-05-01 10:00:00",
        )
        selected_other = _mail("other@example.com", "Other")

        with ExitStack() as stack:
            stack.enter_context(
                patch(
                    "save_mail.get_selected_mail_headers",
                    return_value=[selected_reply, selected_other],
                )
            )
            stack.enter_context(
                patch("save_mail.existing_thread_slug_map", return_value={})
            )
            fetch_candidates = stack.enter_context(
                patch(
                    "save_mail.fetch_conversation_candidates",
                    return_value=[root, selected_reply],
                )
            )
            process_mail = stack.enter_context(
                patch(
                    "save_mail.process_mail_record",
                    side_effect=[
                        {"status": "saved", "slug": "slug-root"},
                        {"status": "saved", "slug": "slug-reply"},
                        {"status": "saved", "slug": "slug-other"},
                    ],
                )
            )
            update_timeline = stack.enter_context(
                patch("save_mail.update_conversation_thread_timeline", return_value=2)
            )

            outcome = save_selected_mail(verbose=False)

        assert outcome["mode"] == "selection"
        assert [r["slug"] for r in outcome["results"]] == [
            "slug-root",
            "slug-reply",
            "slug-other",
        ]
        fetch_candidates.assert_called_once_with(
            selected_reply,
            ["root@example.com"],
            "Project Update",
            include_subject=False,
        )
        assert process_mail.call_count == 3
        update_timeline.assert_called_once()

    def test_multi_selected_mail_skips_subject_fallback_after_exact_miss(
        self,
    ) -> None:
        selected_reply = _mail(
            "reply@example.com",
            "Re: Project Update",
            headers=(
                "Message-ID: <reply@example.com>\n"
                "In-Reply-To: <root@example.com>\n"
                "References: <root@example.com>\n"
            ),
            date="2026-05-01 10:00:00",
        )
        selected_other = _mail("other@example.com", "Other")

        with ExitStack() as stack:
            stack.enter_context(
                patch(
                    "save_mail.get_selected_mail_headers",
                    return_value=[selected_reply, selected_other],
                )
            )
            stack.enter_context(
                patch("save_mail.existing_thread_slug_map", return_value={})
            )
            fetch_candidates = stack.enter_context(
                patch("save_mail.fetch_conversation_candidates", return_value=[])
            )
            process_mail = stack.enter_context(
                patch(
                    "save_mail.process_mail_record",
                    side_effect=[
                        {"status": "saved", "slug": "slug-reply"},
                        {"status": "saved", "slug": "slug-other"},
                    ],
                )
            )

            outcome = save_selected_mail(verbose=False)

        assert outcome["mode"] == "selection"
        assert [r["slug"] for r in outcome["results"]] == [
            "slug-reply",
            "slug-other",
        ]
        fetch_candidates.assert_called_once_with(
            selected_reply,
            ["root@example.com"],
            "Project Update",
            include_subject=False,
        )
        assert process_mail.call_count == 2

    def test_single_mode_with_multi_selection_uses_first_selected_mail(self) -> None:
        first = _mail("first@example.com", "First")
        second = _mail("second@example.com", "Second")

        with ExitStack() as stack:
            stack.enter_context(
                patch(
                    "save_mail.get_selected_mail_headers",
                    return_value=[first, second],
                )
            )
            process_mail = stack.enter_context(
                patch(
                    "save_mail.process_mail_record",
                    return_value={"status": "saved", "slug": "slug-first"},
                )
            )

            outcome = save_selected_mail(single=True, verbose=False)

        assert outcome["mode"] == "single"
        assert outcome["reason"] == "forced"
        process_mail.assert_called_once()
        assert process_mail.call_args.args[0] is first
        assert process_mail.call_args.kwargs["single_mode"] is True

    def test_message_id_match_skips_subject_fallback(self) -> None:
        sent = _mail(
            "a@example.com",
            "Project Update",
            sender="viggomeesters@icloud.com",
            to="person@example.com",
            headers="Message-ID: <a@example.com>\n",
            mailbox_type="sent",
            mailbox="Sent Messages",
            date="2026-05-01 09:00:00",
        )
        selected = _mail(
            "b@example.com",
            "Re: Project Update",
            headers=(
                "Message-ID: <b@example.com>\n"
                "In-Reply-To: <a@example.com>\n"
                "References: <a@example.com>\n"
            ),
            date="2026-05-01 10:00:00",
        )

        with ExitStack() as stack:
            stack.enter_context(
                patch("save_mail.get_selected_mail_headers", return_value=[selected])
            )
            stack.enter_context(
                patch("save_mail.existing_thread_slug_map", return_value={})
            )
            fetch_candidates = stack.enter_context(
                patch("save_mail.fetch_conversation_candidates", return_value=[sent])
            )
            stack.enter_context(
                patch(
                    "save_mail.process_mail_record",
                    return_value={"status": "saved", "slug": "slug"},
                )
            )

            outcome = save_selected_mail(verbose=False)

        assert outcome["mode"] == "conversation"
        fetch_candidates.assert_called_once_with(
            selected,
            ["a@example.com"],
            "Project Update",
            include_subject=False,
        )

    def test_reply_uses_subject_fallback_after_exact_miss(self) -> None:
        sent = _mail(
            "a@example.com",
            "Project Update",
            sender="viggomeesters@icloud.com",
            to="person@example.com",
            headers="Message-ID: <a@example.com>\n",
            mailbox_type="sent",
            mailbox="Sent Messages",
            date="2026-05-01 09:00:00",
        )
        selected = _mail(
            "b@example.com",
            "Re: Project Update",
            headers=(
                "Message-ID: <b@example.com>\n"
                "In-Reply-To: <a@example.com>\n"
                "References: <a@example.com>\n"
            ),
            date="2026-05-01 10:00:00",
        )

        with ExitStack() as stack:
            stack.enter_context(
                patch("save_mail.get_selected_mail_headers", return_value=[selected])
            )
            stack.enter_context(
                patch("save_mail.existing_thread_slug_map", return_value={})
            )
            fetch_candidates = stack.enter_context(
                patch(
                    "save_mail.fetch_conversation_candidates",
                    side_effect=[[], [sent]],
                )
            )
            stack.enter_context(
                patch(
                    "save_mail.process_mail_record",
                    return_value={"status": "saved", "slug": "slug"},
                )
            )

            outcome = save_selected_mail(verbose=False)

        assert outcome["mode"] == "conversation"
        assert fetch_candidates.call_count == 2
        assert fetch_candidates.call_args_list[0].kwargs["include_subject"] is False
        assert fetch_candidates.call_args_list[1].kwargs["include_subject"] is True

    def test_existing_thread_slug_skips_mail_conversation_fetch(self) -> None:
        selected = _mail(
            "b@example.com",
            "Re: Project Update",
            headers=(
                "Message-ID: <b@example.com>\n"
                "In-Reply-To: <a@example.com>\n"
                "References: <a@example.com>\n"
            ),
            date="2026-05-01 10:00:00",
        )

        with ExitStack() as stack:
            stack.enter_context(
                patch("save_mail.get_selected_mail_headers", return_value=[selected])
            )
            stack.enter_context(
                patch(
                    "save_mail.existing_thread_slug_map",
                    return_value={"a@example.com": "slug-a"},
                )
            )
            fetch_candidates = stack.enter_context(
                patch("save_mail.fetch_conversation_candidates")
            )
            process_mail = stack.enter_context(
                patch(
                    "save_mail.process_mail_record",
                    return_value={
                        "status": "saved",
                        "slug": "slug-b",
                        "direction": "received",
                        "path": "/tmp/b.md",
                    },
                )
            )
            update_timeline = stack.enter_context(
                patch("save_mail.update_conversation_thread_timeline", return_value=2)
            )

            outcome = save_selected_mail(verbose=False)

        assert outcome["mode"] == "single"
        assert outcome["reason"] == "existing-thread"
        fetch_candidates.assert_not_called()
        assert process_mail.call_args.kwargs["extra_thread_slugs"] == ["slug-a"]
        assert update_timeline.call_count == 1

    def test_conversation_save_updates_complete_timeline_for_duplicate_and_new(
        self,
    ) -> None:
        root = _mail(
            "a@example.com",
            "Project Update",
            headers="Message-ID: <a@example.com>\n",
            date="2026-05-01 09:00:00",
        )
        sent = _mail(
            "b@example.com",
            "Re: Project Update",
            sender="viggomeesters@icloud.com",
            to="person@example.com",
            headers=(
                "Message-ID: <b@example.com>\n"
                "In-Reply-To: <a@example.com>\n"
                "References: <a@example.com>\n"
            ),
            mailbox_type="sent",
            mailbox="Sent Messages",
            date="2026-05-01 09:30:00",
        )
        selected = _mail(
            "c@example.com",
            "Re: Project Update",
            headers=(
                "Message-ID: <c@example.com>\n"
                "In-Reply-To: <b@example.com>\n"
                "References: <a@example.com> <b@example.com>\n"
            ),
            date="2026-05-01 10:00:00",
        )

        with ExitStack() as stack:
            stack.enter_context(
                patch("save_mail.get_selected_mail_headers", return_value=[selected])
            )
            stack.enter_context(
                patch("save_mail.existing_thread_slug_map", return_value={})
            )
            stack.enter_context(
                patch(
                    "save_mail.fetch_conversation_candidates",
                    return_value=[root, sent, selected],
                )
            )
            stack.enter_context(
                patch(
                    "save_mail.process_mail_record",
                    side_effect=[
                        {
                            "status": "duplicate",
                            "slug": "20260501-0900-mail-root",
                            "direction": "received",
                            "path": "/tmp/root.md",
                        },
                        {
                            "status": "saved",
                            "slug": "20260501-0930-mail-sent",
                            "direction": "sent",
                            "path": "/tmp/sent.md",
                        },
                        {
                            "status": "saved",
                            "slug": "20260501-1000-mail-received",
                            "direction": "received",
                            "path": "/tmp/received.md",
                        },
                    ],
                )
            )
            update_timeline = stack.enter_context(
                patch("save_mail.update_conversation_thread_timeline", return_value=3)
            )

            outcome = save_selected_mail(verbose=False)

        assert outcome["thread_updated"] == 3
        entries = update_timeline.call_args.args[0]
        assert entries == [
            {
                "slug": "20260501-0900-mail-root",
                "direction": "received",
                "path": "/tmp/root.md",
            },
            {
                "slug": "20260501-0930-mail-sent",
                "direction": "sent",
                "path": "/tmp/sent.md",
            },
            {
                "slug": "20260501-1000-mail-received",
                "direction": "received",
                "path": "/tmp/received.md",
            },
        ]


class TestConversationAppleScriptFetch:
    def test_exact_lookup_stays_in_selected_account(self) -> None:
        selected = _mail("b@example.com", "Re: Project Update")
        candidate = _mail("a@example.com", "Project Update")

        with ExitStack() as stack:
            stack.enter_context(
                patch("mail_applescript.is_mail_running", return_value=True)
            )
            fetch_account = stack.enter_context(
                patch(
                    "mail_applescript._fetch_conversation_candidates_from_account",
                    return_value=[candidate],
                )
            )

            result = mail_applescript.fetch_conversation_candidates(
                selected,
                ["a@example.com"],
                "Project Update",
                include_subject=False,
            )

        assert result == [candidate]
        fetch_account.assert_called_once()

    def test_subject_lookup_can_stay_in_selected_account(self) -> None:
        selected = _mail("b@example.com", "Re: Project Update")

        with ExitStack() as stack:
            stack.enter_context(
                patch("mail_applescript.is_mail_running", return_value=True)
            )
            fetch_account = stack.enter_context(
                patch(
                    "mail_applescript._fetch_conversation_candidates_from_account",
                    return_value=[],
                )
            )

            result = mail_applescript.fetch_conversation_candidates(
                selected,
                ["a@example.com"],
                "Project Update",
                include_subject=True,
                search_other_accounts=False,
            )

        assert result == []
        fetch_account.assert_called_once()


class TestIsCalendarInvite:
    """Test is_calendar_invite() detection patterns."""

    def test_standard_subject_prefix(self) -> None:
        assert is_calendar_invite(
            {"subject": "Accepted: Weekly standup", "sender_email": "user@example.com"}
        )

    def test_all_subject_prefixes(self) -> None:
        for prefix in (
            "Invitation:",
            "Accepted:",
            "Declined:",
            "Tentative:",
            "Updated:",
            "Cancelled:",
            "Canceled:",
        ):
            assert is_calendar_invite(
                {"subject": f"{prefix} Meeting", "sender_email": "user@example.com"}
            )

    def test_calendar_sender(self) -> None:
        assert is_calendar_invite(
            {"subject": "Meeting", "sender_email": "noreply@calendar.google.com"}
        )

    def test_ics_attachment(self) -> None:
        assert is_calendar_invite(
            {
                "subject": "1:1 Gilles/Viggo – Simply-4",
                "sender_email": "gilles.van.boven@mccoy-partners.com",
                "att_names": "invite.ics",
            }
        )

    def test_ics_in_mixed_attachments(self) -> None:
        assert is_calendar_invite(
            {
                "subject": "Regular subject",
                "sender_email": "someone@company.com",
                "att_names": "report.pdf, meeting.ICS",
            }
        )

    def test_regular_email_no_match(self) -> None:
        assert not is_calendar_invite(
            {
                "subject": "Regular email",
                "sender_email": "user@example.com",
                "att_names": "report.pdf",
            }
        )

    def test_empty_mail(self) -> None:
        assert not is_calendar_invite({})

    def test_missing_att_names_key(self) -> None:
        assert not is_calendar_invite(
            {"subject": "Hello", "sender_email": "user@example.com"}
        )

    def test_has_calendar_mime_part(self) -> None:
        assert is_calendar_invite(
            {
                "subject": "1:1 Gilles/Viggo",
                "sender_email": "gilles.van.boven@mccoy-partners.com",
                "att_names": "",
                "has_calendar": True,
            }
        )

    def test_has_calendar_false(self) -> None:
        assert not is_calendar_invite(
            {
                "subject": "Regular email",
                "sender_email": "user@example.com",
                "att_names": "",
                "has_calendar": False,
            }
        )


class TestShouldCreateFollowUpTask:
    def test_flagged_inbox_mail_creates_task(self) -> None:
        assert should_create_follow_up_task({"is_flagged": True}, mailbox_type="inbox")

    def test_unflagged_inbox_mail_does_not_create_task(self) -> None:
        assert not should_create_follow_up_task(
            {"is_flagged": False}, mailbox_type="inbox"
        )

    def test_flagged_sent_mail_does_not_create_task(self) -> None:
        assert not should_create_follow_up_task(
            {"is_flagged": True}, mailbox_type="sent"
        )

    def test_force_task_keeps_manual_override(self) -> None:
        assert should_create_follow_up_task(
            {"is_flagged": False}, force_task=True, mailbox_type="sent"
        )


class TestMailSlugCreation:
    def test_make_slug_uses_canonical_month_folder_for_collisions(
        self, tmp_path, monkeypatch
    ) -> None:
        vault = tmp_path / "vault"
        notes = vault / "10_notes"
        monkeypatch.setattr(save_mail.brain_lib.cfg, "vault_root", vault)
        monkeypatch.setattr(save_mail.brain_lib.cfg, "vault_notes", notes)
        month = notes / "2026-05"
        month.mkdir(parents=True)
        (month / "20260516-1345-mail-anne-project-update.md").write_text(
            "", encoding="utf-8"
        )

        slug, ts = save_mail.make_slug(
            datetime(2026, 5, 16, 13, 45), "anne", "Project Update"
        )

        assert ts == "20260516-1345"
        assert slug == "20260516-1345-mail-anne-project-update-2"

    def test_display_only_outlook_sender_is_used_for_entity_slug(
        self, tmp_path, monkeypatch
    ) -> None:
        vault = tmp_path / "vault"
        notes = vault / "10_notes"
        entities = vault / "system" / "entities"
        monkeypatch.setattr(save_mail, "VAULT_ROOT", vault)
        monkeypatch.setattr(save_mail.brain_lib.cfg, "vault_root", vault)
        monkeypatch.setattr(save_mail.brain_lib.cfg, "vault_notes", notes)
        monkeypatch.setattr(save_mail.brain_lib.cfg, "vault_entities", entities)
        monkeypatch.setattr(save_mail.brain_lib, "link_note_in_daily", lambda *a, **k: None)

        result = save_mail.create_mail_note(
            {
                "subject": "AskIT: Je aanvraag REQ0116499",
                "sender_email": "askit servicenow",
                "sender_display": "AskIT ServiceNow",
                "date_str": "2026-06-17 11:01:00",
                "message_id": "outlook-display-only@example.local",
                "account": "viggo.meesters@bam.com",
                "mailbox": "Inbox",
                "to": "",
                "cc": "",
                "att_count": "0",
                "att_names": "",
                "is_flagged": False,
                "all_headers": "",
                "mailbox_type": "inbox",
                "body": "Ticket geregistreerd.",
                "mail_client": "outlook",
            }
        )

        assert result["entity"] == "askit-servicenow"
        assert result["entity_source"] == "outlook-display-name"
        assert result["entity_confidence"] == 0.65
        assert result["slug"] == "20260617-1101-mail-askit-servicenow-askit-je-aanvraag"
        content = Path(result["path"]).read_text(encoding="utf-8")
        assert 'entity: ["askit-servicenow"]' in content
        assert "entity_source: outlook-display-name" in content
        assert "entity_confidence: 0.65" in content
        assert "to: viggo.meesters@bam.com" in content
        assert "📧 From: [[askit-servicenow]] (askit servicenow)" in content
        assert "📬 To: viggo.meesters@bam.com" in content
        entity_content = (entities / "askit-servicenow.md").read_text(encoding="utf-8")
        assert "source: auto-created-from-outlook-display-name" in entity_content
        assert "title: AskIT ServiceNow" in entity_content


class TestEntityResolutionMetadata:
    def test_known_email_map_records_exact_source(self) -> None:
        result = save_mail.resolve_entity_details("noreply@email.openai.com")

        assert result.slug == "openai"
        assert result.direction == "received"
        assert result.source == "known-email-map"
        assert result.confidence == 1.0

    def test_self_sent_mail_uses_recipient_entity_source(self) -> None:
        result = save_mail.resolve_entity_details(
            "viggomeesters@icloud.com",
            "person@example.com",
        )

        assert result.slug == "example"
        assert result.direction == "sent"
        assert result.source.startswith("sent-recipient:")
        assert result.confidence <= 0.95

    def test_relay_map_records_relay_source(self) -> None:
        result = save_mail.resolve_entity_details(
            "anything-instant-gaming@privaterelay.appleid.com"
        )

        assert result.slug == "instant-gaming"
        assert result.source == "apple-relay-map"
        assert result.confidence == 0.9

    def test_missing_email_records_unknown_source(self) -> None:
        result = save_mail.resolve_entity_details("AskIT ServiceNow")

        assert result.slug == "unknown-sender"
        assert result.source == "missing-email"
        assert result.confidence == 0.0


class TestProjectTopicRoutingMetadata:
    def test_mail_note_records_project_slug_code_and_topic_sources(
        self, tmp_path, monkeypatch
    ) -> None:
        vault = tmp_path / "vault"
        notes = vault / "10_notes"
        monkeypatch.setattr(save_mail.brain_lib.cfg, "vault_root", vault)
        monkeypatch.setattr(save_mail.brain_lib.cfg, "vault_notes", notes)
        monkeypatch.setattr(save_mail.brain_lib, "link_note_in_daily", lambda *a, **k: None)
        monkeypatch.setattr(
            save_mail.brain_lib,
            "resolve_project_code",
            lambda slug: "mailcap" if slug == "2026-06-mail-capture" else slug,
        )
        monkeypatch.setattr(save_mail, "suggest_topics", lambda *a, **k: ["obsidian"])
        monkeypatch.setattr(
            save_mail,
            "detect_mail_project",
            lambda mail: mail_project_rules.MailProjectMatch(
                project="2026-06-mail-capture",
                area="self",
                source="mail_code",
                confidence=0.9,
                matched="mailcap",
            ),
        )

        result = save_mail.create_mail_note(
            {
                "subject": "Project Update",
                "sender_email": "person@example.com",
                "sender_display": "Person Example",
                "date_str": "2026-05-01 10:00:00",
                "message_id": "routing@example.com",
                "account": "iCloud",
                "mailbox": "Inbox",
                "to": "viggomeesters@icloud.com",
                "cc": "",
                "att_count": "0",
                "att_names": "",
                "is_flagged": False,
                "all_headers": "",
                "mailbox_type": "inbox",
                "body": "Saved body.",
            }
        )

        assert result["project"] == "mailcap"
        assert result["project_slug"] == "2026-06-mail-capture"
        assert result["project_source"] == "mail_code"
        assert result["project_confidence"] == 0.9
        assert result["topics"] == ["obsidian"]
        assert result["topics_source"] == "entity-or-subject"
        assert result["topics_confidence"] == 0.7

        content = Path(result["path"]).read_text(encoding="utf-8")
        assert "project: mailcap" in content
        assert "project_slug: 2026-06-mail-capture" in content
        assert "project_source: mail_code" in content
        assert "topics_source: entity-or-subject" in content
        assert "topics_confidence: 0.7" in content


class TestFollowUpTaskCreation:
    def test_uses_task_prefix(self, tmp_path, monkeypatch) -> None:
        class FixedDateTime(datetime):
            @classmethod
            def now(cls):
                return cls(2026, 5, 16, 13, 45)

        monkeypatch.setattr(save_mail, "datetime", FixedDateTime)
        vault = tmp_path / "vault"
        notes = vault / "10_notes"
        monkeypatch.setattr(save_mail.brain_lib.cfg, "vault_root", vault)
        monkeypatch.setattr(save_mail.brain_lib.cfg, "vault_notes", notes)

        result = save_mail.create_follow_up_task(
            subject="Re: Project Update",
            entity_slug="anne",
            mail_slug="20260516-1300-mail-anne-project-update",
            area="self",
            project=None,
        )

        assert result["slug"] == "20260516-1345-task-follow-up-project-update"
        path = notes / "2026-05" / "20260516-1345-task-follow-up-project-update.md"
        assert Path(result["path"]) == path
        assert path.exists()
        assert (
            result["obsidian_file"]
            == "10_notes/2026-05/20260516-1345-task-follow-up-project-update"
        )


class TestSavedDataSize:
    def test_format_data_size_uses_human_units(self) -> None:
        assert save_mail.format_data_size(512) == "512 B"
        assert save_mail.format_data_size(1536) == "1.5 KB"
        assert save_mail.format_data_size(2 * 1024 * 1024) == "2.0 MB"

    def test_create_mail_note_returns_note_size(self, tmp_path, monkeypatch) -> None:
        vault = tmp_path / "vault"
        notes = vault / "10_notes"
        monkeypatch.setattr(save_mail.brain_lib.cfg, "vault_root", vault)
        monkeypatch.setattr(save_mail.brain_lib.cfg, "vault_notes", notes)
        monkeypatch.setattr(
            save_mail.brain_lib,
            "link_note_in_daily",
            lambda *a, **k: None,
        )

        result = save_mail.create_mail_note(
            {
                "subject": "Re: Project Update",
                "sender_email": "person@example.com",
                "sender_display": "Person Example",
                "date_str": "2026-05-01 10:00:00",
                "message_id": "size@example.com",
                "account": "iCloud",
                "mailbox": "Inbox",
                "to": "viggomeesters@icloud.com",
                "cc": "",
                "att_count": "0",
                "att_names": "",
                "is_flagged": False,
                "all_headers": "",
                "mailbox_type": "inbox",
                "body": "Saved body.",
            }
        )

        path = Path(result["path"])
        assert result["note_bytes"] == path.stat().st_size
        assert result["note_size"] == save_mail.format_data_size(path.stat().st_size)
        assert result["capture_source"] == "save-mail"
        assert result["capture_version"] == 1
        assert result["enrichment_status"] == "pending"
        assert result["enrichment_version"] == 0
        assert result["raw_subject"] == "Re: Project Update"
        assert result["clean_subject"] == "Project Update"
        assert result["sender_domain"] == "example.com"

        content = path.read_text(encoding="utf-8")
        assert "capture_source: save-mail" in content
        assert "capture_version: 1" in content
        assert "enrichment_status: pending" in content
        assert "enrichment_version: 0" in content
        assert 'raw_subject: "Re: Project Update"' in content
        assert "clean_subject: Project Update" in content
        assert "sender_domain: example.com" in content


class TestArchiveRouting:
    def test_display_only_sender_gets_stable_entity_slug(self) -> None:
        assert (
            save_mail._entity_slug_from_display_name("askit servicenow")
            == "askit-servicenow"
        )

    def test_unknown_display_sender_does_not_create_entity_slug(self) -> None:
        assert save_mail._entity_slug_from_display_name("unknown sender") is None

    def test_inbox_mail_archives_after_save(self) -> None:
        mail = _mail("one@example.com", "Inbox mail")

        assert save_mail._should_archive_after_save(
            mail, no_archive=False, single_mode=True
        )

    def test_outlook_inbox_mail_archives_after_save(self) -> None:
        mail = _mail("outlook-id", "Outlook inbox")
        mail["mail_client"] = "outlook"

        assert save_mail._should_archive_after_save(
            mail, no_archive=False, single_mode=True
        )

    def test_sent_mail_does_not_archive_after_selected_save(self) -> None:
        mail = _mail(
            "sent@example.com",
            "Sent mail",
            sender="viggomeesters@icloud.com",
            to="person@example.com",
            mailbox_type="sent",
            mailbox="Sent Messages",
        )

        assert not save_mail._should_archive_after_save(
            mail, no_archive=False, single_mode=True
        )

    def test_deleted_mail_does_not_archive_after_selected_save(self) -> None:
        mail = _mail(
            "deleted@example.com",
            "Accepted: Planning",
            mailbox_type="deleted",
            mailbox="Deleted Messages",
        )

        assert not save_mail._should_archive_after_save(
            mail, no_archive=False, single_mode=True
        )

    def test_archive_failure_does_not_fail_saved_mail(self) -> None:
        mail = _mail("one@example.com", "Budget Thuis in een nieuw jasje")

        with ExitStack() as stack:
            stack.enter_context(patch("save_mail.is_duplicate", return_value=None))
            stack.enter_context(patch("save_mail.fetch_mail_body", return_value="body"))
            stack.enter_context(
                patch(
                    "save_mail.create_mail_note",
                    return_value={
                        "slug": "20260509-0956-mail-budget-thuis",
                        "path": "/tmp/budget.md",
                    },
                )
            )
            stack.enter_context(
                patch(
                    "save_mail.archive_mail",
                    return_value="ERROR:AppleScript timed out after 15s",
                )
            )
            stack.enter_context(patch("save_mail._write_log"))

            result = save_mail.process_mail_record(mail, single_mode=True, verbose=False)

        assert result["status"] == "saved"
        assert result["archived"] == "ERROR:AppleScript timed out after 15s"
        assert result["archive_error"] == "ERROR:AppleScript timed out after 15s"


class TestOutlookArchive:
    def test_new_outlook_archive_uses_menu_after_matching_selection(self) -> None:
        calls: list[str] = []

        def fake_run_applescript(script: str, timeout: int = 10) -> str:
            calls.append(script)
            return "OK"

        with ExitStack() as stack:
            stack.enter_context(
                patch("mail_outlook_applescript._is_new_outlook_enabled", return_value=True)
            )
            stack.enter_context(
                patch(
                    "mail_outlook_applescript._get_selected_mail_header_from_accessibility",
                    return_value={"message_id": "outlook-id"},
                )
            )
            stack.enter_context(
                patch(
                    "mail_outlook_applescript.run_applescript",
                    side_effect=fake_run_applescript,
                )
            )

            result = mail_outlook_applescript.archive_mail("outlook-id", "account")

        assert result == "OK"
        assert calls
        assert "Archive" in calls[0]

    def test_new_outlook_archive_skips_when_selection_changed(self) -> None:
        with ExitStack() as stack:
            stack.enter_context(
                patch("mail_outlook_applescript._is_new_outlook_enabled", return_value=True)
            )
            stack.enter_context(
                patch(
                    "mail_outlook_applescript._get_selected_mail_header_from_accessibility",
                    return_value={"message_id": "other-id"},
                )
            )
            run_applescript = stack.enter_context(
                patch("mail_outlook_applescript.run_applescript")
            )

            result = mail_outlook_applescript.archive_mail("outlook-id", "account")

        assert result == "SKIPPED:outlook selection changed"
        run_applescript.assert_not_called()

    def test_synthetic_outlook_id_requires_selection_check(self) -> None:
        with ExitStack() as stack:
            stack.enter_context(
                patch("mail_outlook_applescript._is_new_outlook_enabled", return_value=False)
            )
            stack.enter_context(
                patch(
                    "mail_outlook_applescript._get_selected_mail_header_from_accessibility",
                    return_value={"message_id": "other-id"},
                )
            )
            run_applescript = stack.enter_context(
                patch("mail_outlook_applescript.run_applescript")
            )

            result = mail_outlook_applescript.archive_mail(
                "outlook-abc123@local.outlook",
                "account",
            )

        assert result == "SKIPPED:outlook selection changed"
        run_applescript.assert_not_called()
