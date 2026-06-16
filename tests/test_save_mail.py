"""Tests for save_mail.py calendar invite detection."""

from __future__ import annotations

import sys
import sqlite3
from types import SimpleNamespace
from contextlib import ExitStack
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "shared"))

import mail_applescript
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


class TestConversationSaveFlow:
    def test_explicit_outlook_client_routes_through_outlook_adapter(self) -> None:
        selected = _mail("outlook-id", "Outlook subject")
        selected.pop("mail_client", None)
        adapter_calls: list[str] = []
        created: list[dict] = []

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


class TestArchiveRouting:
    def test_inbox_mail_archives_after_save(self) -> None:
        mail = _mail("one@example.com", "Inbox mail")

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
