"""Offline checks for outbound Teams routing and Word timeout cleanup."""

import asyncio
import contextlib
import importlib.util
import io
import json
import subprocess
import sys
import unittest
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import AsyncMock, Mock, patch


ROOT = Path(__file__).resolve().parents[1]


def load_script(name, relative_path):
    path = ROOT / relative_path
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class SafeDispatchTests(unittest.TestCase):
    def test_safety_hooks_deny_malformed_input(self):
        for path in (
            ".agents/plugins/outlook-tools/hooks/safety_check.py",
            ".agents/plugins/teams-tools/hooks/scripts/teams_safety_gate.py",
        ):
            module = load_script("safety_hook", path)
            for raw in ("", "{invalid", "{}"):
                with self.subTest(path=path, raw=raw):
                    output = io.StringIO()
                    with patch("sys.stdin", io.StringIO(raw)), contextlib.redirect_stdout(output):
                        module.main()
                    self.assertEqual(json.loads(output.getvalue())["decision"], "deny")

    def test_word_timeout_kills_only_owned_pid(self):
        module = load_script(
            "docx_to_pdf_safe",
            ".agents/plugins/docx-tools/skills/docx-to-pdf/scripts/docx_to_pdf.py",
        )
        timeout = subprocess.TimeoutExpired("word", 1, output=b"WORD_PID=123\r\n")
        runner = Mock(side_effect=[timeout, subprocess.CompletedProcess([], 0)])
        with patch.object(module.platform, "system", return_value="Windows"), patch.object(
            module.subprocess, "run", runner
        ):
            self.assertFalse(module.convert_with_word(Path("in.docx"), Path("out.pdf")))
        cleanup_command = runner.call_args_list[1].args[0]
        self.assertEqual(Path(cleanup_command[0]).name.lower(), "taskkill.exe")
        self.assertEqual(cleanup_command[1:], ["/F", "/PID", "123"])

    def test_word_timeout_without_pid_does_not_kill_other_processes(self):
        module = load_script(
            "docx_to_pdf_no_pid",
            ".agents/plugins/docx-tools/skills/docx-to-pdf/scripts/docx_to_pdf.py",
        )
        runner = Mock(side_effect=subprocess.TimeoutExpired("word", 1))
        with patch.object(module.platform, "system", return_value="Windows"), patch.object(
            module.subprocess, "run", runner
        ):
            self.assertFalse(module.convert_with_word(Path("in.docx"), Path("out.pdf")))
        runner.assert_called_once()

    def test_word_worker_quits_after_document_close_fails(self):
        module = load_script(
            "docx_to_pdf_cleanup",
            ".agents/plugins/docx-tools/skills/docx-to-pdf/scripts/docx_to_pdf.py",
        )
        with patch.object(module.platform, "system", return_value="Windows"), patch.object(
            module.subprocess, "run", return_value=subprocess.CompletedProcess([], 1, "", "")
        ) as runner, contextlib.redirect_stderr(io.StringIO()):
            self.assertFalse(module.convert_with_word(Path("in.docx"), Path("out.pdf")))
        worker_code = runner.call_args.args[0][2]
        document = Mock()
        document.Close.side_effect = RuntimeError("close failed")
        word = Mock(Hwnd=1)
        word.Documents.Open.return_value = document
        com_client = Mock()
        com_client.DispatchEx.return_value = word
        pythoncom = Mock()
        win32process = Mock()
        win32process.GetWindowThreadProcessId.return_value = (0, 123)
        with patch.dict(sys.modules, {"win32com": Mock(client=com_client), "win32com.client": com_client,
                                      "pythoncom": pythoncom, "win32process": win32process}), patch.object(
            sys, "argv", ["worker", "in.docx", "out.pdf"]
        ), contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit):
                exec(worker_code, {"__name__": "__main__"})
        word.Quit.assert_called_once()
        pythoncom.CoUninitialize.assert_called_once()

    def test_recipient_selection_resolves_chat(self):
        shared = ROOT / ".agents/plugins/teams-tools/shared"
        sys.path.insert(0, str(shared))
        try:
            module = load_script(
                "message_send_safe",
                ".agents/plugins/teams-tools/skills/teams-chat-writer/scripts/message_send.py",
            )
            # Explicit chat ID takes precedence
            self.assertEqual(
                asyncio.run(module.resolve_recipient("chat-123", None)),
                ("chat-123", "chat-123"),
            )
            self.assertEqual(
                asyncio.run(module.resolve_recipient("chat-123", "Alice")),
                ("chat-123", "chat-123"),
            )

            # Active chat fallback when no contact or chat_id specified
            with patch.object(module.teams_client, "get_active_target", return_value={"title": "Team Channel"}):
                self.assertEqual(
                    asyncio.run(module.resolve_recipient(None, None)),
                    ("active", "Team Channel"),
                )

            # Contact resolution via list_chats
            chats = [{"id": "chat-alice", "title": "Alice Smith"}]
            with patch.object(module.teams_client, "list_chats", new_callable=AsyncMock, return_value=chats):
                self.assertEqual(
                    asyncio.run(module.resolve_recipient(None, "Alice")),
                    ("chat-alice", "Alice Smith"),
                )

            # Error when no chat matches contact
            with patch.object(module.teams_client, "list_chats", new_callable=AsyncMock, return_value=[]):
                with self.assertRaisesRegex(ValueError, "No chat found"):
                    asyncio.run(module.resolve_recipient(None, "Unknown Contact"))
        finally:
            sys.path.remove(str(shared))

    def test_outlook_send_recipients_validation(self):
        module = load_script(
            "mail_compose_safe", ".agents/plugins/outlook-tools/skills/outlook-mail/scripts/mail_compose.py"
        )
        entry = Mock(Type="SMTP", Address="user@example.com")
        recipient = Mock(Type=1, AddressEntry=entry)
        recipients = Mock(Count=1, ResolveAll=Mock(return_value=True), Item=Mock(return_value=recipient))
        mail = Mock(Recipients=recipients)

        # Valid SMTP address passes
        module.validate_send_recipients(mail)

        # Multiple recipients pass
        recipients.Count = 2
        module.validate_send_recipients(mail)
        recipients.Count = 1

        # Valid Exchange address passes
        ex_user = Mock(PrimarySmtpAddress="exchange.user@company.com")
        ex_entry = Mock(Type="EX", GetExchangeUser=Mock(return_value=ex_user))
        recipient.AddressEntry = ex_entry
        module.validate_send_recipients(mail)
        recipient.AddressEntry = entry

        # Empty / whitespace addresses fail
        for invalid_address in ("", "   ", None):
            entry.Address = invalid_address
            with self.assertRaises(ValueError):
                module.validate_send_recipients(mail)
        entry.Address = "user@example.com"

        # Missing AddressEntry fails
        recipient.AddressEntry = None
        with self.assertRaises(ValueError):
            module.validate_send_recipients(mail)
        recipient.AddressEntry = entry

        # Unresolved recipients fail
        recipients.ResolveAll.return_value = False
        with self.assertRaises(ValueError):
            module.validate_send_recipients(mail)
        recipients.ResolveAll.return_value = True

        # Zero recipients fail
        recipients.Count = 0
        with self.assertRaises(ValueError):
            module.validate_send_recipients(mail)

    def test_calendar_invitation_recipients_validation(self):
        try:
            module = load_script(
                "calendar_guard_safe", ".agents/plugins/outlook-tools/skills/outlook-calendar/scripts/calendar_manage.py"
            )
        except ImportError as error:
            self.skipTest(f"Calendar dependencies unavailable: {error}")
        entry = Mock(Type="SMTP", Address="attendee@example.com")
        recipient = Mock(Type=1, AddressEntry=entry)
        recipients = Mock(Count=1, ResolveAll=Mock(return_value=True), Item=Mock(return_value=recipient))
        appointment = Mock(Recipients=recipients)

        # Valid attendee passes
        module.validate_invitation_recipients(appointment)

        # Multiple attendees pass
        recipients.Count = 3
        module.validate_invitation_recipients(appointment)
        recipients.Count = 1

        # Valid Exchange attendee passes
        ex_user = Mock(PrimarySmtpAddress="exchange.attendee@company.com")
        ex_entry = Mock(Type="EX", GetExchangeUser=Mock(return_value=ex_user))
        recipient.AddressEntry = ex_entry
        module.validate_invitation_recipients(appointment)
        recipient.AddressEntry = entry

        # Empty / whitespace addresses fail
        for invalid_address in ("", "   ", None):
            entry.Address = invalid_address
            with self.assertRaises(ValueError):
                module.validate_invitation_recipients(appointment)
        entry.Address = "attendee@example.com"

        # Missing AddressEntry fails
        recipient.AddressEntry = None
        with self.assertRaises(ValueError):
            module.validate_invitation_recipients(appointment)
        recipient.AddressEntry = entry

        # Unresolved recipients fail
        recipients.ResolveAll.return_value = False
        with self.assertRaises(ValueError):
            module.validate_invitation_recipients(appointment)
        recipients.ResolveAll.return_value = True

        # Zero attendees fail
        recipients.Count = 0
        with self.assertRaises(ValueError):
            module.validate_invitation_recipients(appointment)

    def test_mail_search_restrict_failure_still_filters_before_delete(self):
        module = load_script(
            "mail_search_fallback", ".agents/plugins/outlook-tools/skills/outlook-mail/scripts/mail_search.py"
        )
        old_mail = Mock(Class=43, Subject="Old", SenderName="Sender", SenderEmailAddress="sender@example.com",
                        Body="", Attachments=Mock(Count=0), UnRead=True, ReceivedTime="2020-01-01 12:00:00",
                        EntryID="old")
        recent_mail = Mock(Class=43, Subject="Recent", SenderName="Sender", SenderEmailAddress="sender@example.com",
                           Body="", Attachments=Mock(Count=0), UnRead=True, ReceivedTime="2025-01-02 12:00:00",
                           EntryID="recent")
        items = Mock()
        items.Restrict.side_effect = RuntimeError("restriction unavailable")
        items.__iter__ = Mock(return_value=iter([old_mail, recent_mail]))
        folder = Mock(Items=items, Name="Inbox")
        namespace = Mock()
        namespace.GetDefaultFolder.return_value = folder
        namespace.GetItemFromID.return_value = recent_mail
        application = Mock()
        application.GetNamespace.return_value = namespace
        com_client = Mock()
        com_client.GetActiveObject.return_value = application
        win32com = Mock(client=com_client)
        pythoncom = Mock()
        with patch.dict(sys.modules, {"win32com": win32com, "win32com.client": com_client,
                                      "pythoncom": pythoncom}), patch.object(
            sys, "argv", ["mail_search.py", "--since", "2025-01-01", "--delete", "--apply", "--confirm-delete-ids", "recent", "--format", "json"]
        ), contextlib.redirect_stdout(io.StringIO()) as output, contextlib.redirect_stderr(io.StringIO()):
            module.main()
        self.assertEqual(json.loads(output.getvalue())["deleted_count"], 1)
        namespace.GetItemFromID.assert_called_once_with("recent")

    def test_outlook_dry_run_never_connects_to_com(self):
        scripts = (
            ("mail_compose_preview", ".agents/plugins/outlook-tools/skills/outlook-mail/scripts/mail_compose.py",
             ["mail_compose.py", "--send", "--format", "json"]),
            ("calendar_preview", ".agents/plugins/outlook-tools/skills/outlook-calendar/scripts/calendar_manage.py",
             ["calendar_manage.py", "--cancel", "--id", "event", "--format", "json"]),
        )
        com_client = Mock()
        pythoncom = Mock()
        with patch.dict(sys.modules, {"win32com": Mock(client=com_client), "win32com.client": com_client,
                                      "pythoncom": pythoncom}):
            for name, path, argv in scripts:
                with self.subTest(name=name):
                    module = load_script(name, path)
                    with patch.object(sys, "argv", argv), contextlib.redirect_stdout(io.StringIO()) as output:
                        module.main()
                    self.assertEqual(json.loads(output.getvalue())["status"], "dry_run")
                    pythoncom.CoInitialize.assert_not_called()
                    com_client.GetActiveObject.assert_not_called()
                    com_client.Dispatch.assert_not_called()

    def test_calendar_cancel_rejects_mismatched_id_before_com(self):
        module = load_script(
            "calendar_cancel_guard", ".agents/plugins/outlook-tools/skills/outlook-calendar/scripts/calendar_manage.py"
        )
        com_client = Mock()
        pythoncom = Mock()
        with patch.dict(sys.modules, {"win32com": Mock(client=com_client), "win32com.client": com_client,
                                      "pythoncom": pythoncom}), patch.object(
            sys, "argv", ["calendar_manage.py", "--cancel", "--id", "event", "--confirm-cancel-id", "other",
                         "--apply", "--format", "json"]
        ), contextlib.redirect_stdout(io.StringIO()) as output, contextlib.redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit):
                module.main()
        self.assertEqual(json.loads(output.getvalue())["status"], "error")
        pythoncom.CoInitialize.assert_not_called()
        com_client.GetActiveObject.assert_not_called()
        com_client.Dispatch.assert_not_called()

    def test_mail_delete_requires_apply_and_matching_ids(self):
        module = load_script("mail_delete_guard", ".agents/plugins/outlook-tools/skills/outlook-mail/scripts/mail_search.py")
        mail = Mock(Class=43, Subject="Subject", SenderName="Sender", SenderEmailAddress="sender@example.com",
                    Body="", Attachments=Mock(Count=0), UnRead=True, ReceivedTime="2025-01-02 12:00:00",
                    EntryID="current")
        items = Mock()
        items.__iter__ = Mock(side_effect=lambda: iter([mail]))
        folder = Mock(Items=items, Name="Inbox")
        namespace = Mock()
        namespace.GetDefaultFolder.return_value = folder
        application = Mock()
        application.GetNamespace.return_value = namespace
        com_client = Mock()
        com_client.GetActiveObject.return_value = application
        with patch.dict(sys.modules, {"win32com": Mock(client=com_client), "win32com.client": com_client,
                                      "pythoncom": Mock()}):
            for flags, expected_status in (([], "dry_run"), (["--apply", "--confirm-delete-ids", "stale"], "error")):
                with self.subTest(flags=flags), patch.object(sys, "argv", ["mail_search.py", "--delete", "--format", "json", *flags]), contextlib.redirect_stdout(io.StringIO()) as output:
                    try:
                        module.main()
                    except SystemExit as error:
                        self.assertEqual(error.code, 1)
                    self.assertEqual(json.loads(output.getvalue())["status"], expected_status)
                mail.Delete.assert_not_called()
                namespace.GetItemFromID.assert_not_called()

    def test_docx_editor_dry_run_leaves_document_and_backup_unchanged(self):
        module = load_script("docx_editor_preview", ".agents/plugins/docx-tools/skills/docx-editor/scripts/docx_edit.py")
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "fixture.docx"
            source.write_bytes(b"original")
            with patch.object(sys, "argv", ["docx_edit.py", str(source), "append-paragraph", "--text", "New"]), patch.object(
                module, "create_backup"
            ) as backup, patch.object(module, "save_document_safely") as save, contextlib.redirect_stdout(io.StringIO()):
                module.main()
            backup.assert_not_called()
            save.assert_not_called()
            self.assertEqual(source.read_bytes(), b"original")

    def test_docx_editor_applies_changes_with_trailing_apply_flag(self):
        module = load_script("docx_editor_apply", ".agents/plugins/docx-tools/skills/docx-editor/scripts/docx_edit.py")
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "fixture.docx"
            source.write_bytes(b"original")
            with patch.object(sys, "argv", ["docx_edit.py", str(source), "append-paragraph", "--text", "New", "--apply"]), patch.object(
                module, "create_backup"
            ) as backup, patch.object(module, "save_document_safely") as save, patch.object(
                module, "Document"
            ), contextlib.redirect_stdout(io.StringIO()):
                module.main()
            backup.assert_called_once_with(source.resolve())
            save.assert_called_once()

    def test_teams_download_keeps_existing_file(self):
        module = load_script("teams_download_collision", ".agents/plugins/teams-tools/shared/teams_client.py")
        with TemporaryDirectory() as temp_dir:
            destination = Path(temp_dir) / "existing.bin"
            destination.write_bytes(b"original")
            with patch.object(module, "evaluate_js", new_callable=AsyncMock) as evaluate:
                with self.assertRaises(FileExistsError):
                    asyncio.run(module.download_attachment_to_file("https://teams.microsoft.com/file", str(destination)))
                evaluate.assert_not_called()
            self.assertEqual(destination.read_bytes(), b"original")

    def test_teams_download_rejects_invalid_payload_without_writing(self):
        module = load_script("teams_download_limit", ".agents/plugins/teams-tools/shared/teams_client.py")
        with TemporaryDirectory() as temp_dir:
            destination = Path(temp_dir) / "attachment.bin"
            for payload in ({"success": True, "size": 20 * 1024 * 1024 + 1, "base64": "AA=="},
                            {"success": True, "size": 1, "base64": "!"}):
                with self.subTest(payload=payload), patch.object(
                    module, "evaluate_js", new_callable=AsyncMock, return_value=payload
                ):
                    with self.assertRaises(RuntimeError):
                        asyncio.run(module.download_attachment_to_file("https://teams.microsoft.com/file", str(destination)))
                    self.assertFalse(destination.exists())

    def test_teams_send_rejects_unknown_cdp_result(self):
        module = load_script("teams_client_unknown", ".agents/plugins/teams-tools/shared/teams_client.py")
        with patch.object(module, "evaluate_js", new_callable=AsyncMock, return_value=None):
            with self.assertRaisesRegex(RuntimeError, "Unexpected CDP response"):
                asyncio.run(module.send_message_to_active_chat("test", "chat-123"))

    def test_send_guard_is_before_click(self):
        module = load_script(
            "teams_client_safe", ".agents/plugins/teams-tools/shared/teams_client.py"
        )
        with patch.object(module, "evaluate_js", new_callable=AsyncMock, return_value={"success": False, "error": "Active chat does not match requested chat ID."}) as evaluate:
            with self.assertRaisesRegex(RuntimeError, "Active chat does not match"):
                asyncio.run(module.send_message_to_active_chat("test", "chat-123"))
        javascript = evaluate.call_args.args[0]
        self.assertLess(javascript.index("Active chat does not match"), javascript.index("sendBtn.click()"))
        self.assertIn('"chat-123"', javascript)

    def test_download_rejects_untrusted_url_before_browser_access(self):
        module = load_script(
            "teams_client_download_safe", ".agents/plugins/teams-tools/shared/teams_client.py"
        )
        with patch.object(module, "evaluate_js", new_callable=AsyncMock) as evaluate:
            with self.assertRaisesRegex(ValueError, "trusted Microsoft"):
                asyncio.run(module.download_attachment_to_file("https://example.org/file", "unused"))
        evaluate.assert_not_called()

    def test_docx_comment_parser_rejects_entities(self):
        try:
            module = load_script(
                "docx_to_md_safe",
                ".agents/plugins/docx-tools/skills/docx-to-markdown/scripts/docx_to_md.py",
            )
        except ImportError as error:
            self.skipTest(f"DOCX dependencies unavailable: {error}")
        with TemporaryDirectory() as temp_dir:
            file_path = Path(temp_dir) / "untrusted.docx"
            with zipfile.ZipFile(file_path, "w") as archive:
                archive.writestr("word/comments.xml", '<!DOCTYPE x [<!ENTITY x "expanded">]><root>&x;</root>')
            self.assertEqual(module.extract_comments_from_docx(file_path), [])

    def test_recalc_without_engine_reports_failure(self):
        module = load_script("excel_recalc_safe", ".agents/plugins/excel-tools/skills/xlsx/scripts/recalc.py")
        with TemporaryDirectory() as temp_dir:
            workbook = Path(temp_dir) / "fixture.xlsx"
            workbook.write_bytes(b"fixture")
            with patch.object(module, "is_excel_com_available", return_value=False), patch.object(
                module, "find_libreoffice_binary", return_value=None
            ), patch.object(module, "check_external_links_at_risk", return_value=[]):
                self.assertIn("error", module.recalculate(workbook))
            self.assertEqual(workbook.read_bytes(), b"fixture")

    def test_docx_backup_collision_keeps_existing_copy(self):
        try:
            module = load_script(
                "docx_editor_safe", ".agents/plugins/docx-tools/skills/docx-editor/scripts/docx_edit.py"
            )
        except ImportError as error:
            self.skipTest(f"DOCX dependencies unavailable: {error}")
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "fixture.docx"
            source.write_bytes(b"original")
            backup = module.create_backup(source)
            with patch.object(module.sys, "stderr"):
                with self.assertRaises(SystemExit):
                    module.create_backup(source)
            self.assertEqual(backup.read_bytes(), b"original")


if __name__ == "__main__":
    unittest.main()