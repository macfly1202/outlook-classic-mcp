"""Tests for draft, conversation, search, and send-guard mail helpers."""

from __future__ import annotations

import datetime as dt
import sys
import types
from types import SimpleNamespace

import pytest

sys.modules.setdefault("pythoncom", types.ModuleType("pythoncom"))

from outlook_mcp.client import mail as mail_client
from outlook_mcp.constants import OL_CLASS_MAIL
from outlook_mcp.errors import OutlookError


class FakeAttachments(list):
    @property
    def Count(self) -> int:
        return len(self)

    def Add(self, path: str):
        self.append(SimpleNamespace(FileName=path.rsplit("/", 1)[-1], Size=1))

    def Remove(self, index: int) -> None:
        del self[index - 1]


class FakeItems(list):
    def Sort(self, *_args, **_kwargs) -> None:
        self.sort(
            key=lambda item: item._timestamp or dt.datetime.min,
            reverse=True,
        )

    def Restrict(self, _query: str):
        return list(self)


class FakeFolder:
    def __init__(self, name: str, parent: "FakeFolder | None" = None):
        self.Name = name
        self.Parent = parent
        self.Folders: list[FakeFolder] = []
        self.Items = FakeItems()
        if parent is None:
            self.FolderPath = f"\\\\{name}"
        else:
            self.FolderPath = f"{parent.FolderPath}\\{name}"
            parent.Folders.append(self)


class FakeMailItem:
    def __init__(
        self,
        entry_id: str,
        subject: str,
        body: str,
        folder: FakeFolder,
        *,
        conversation_id: str = "conv-1",
        sender_name: str = "Alice",
        sender_address: str = "alice@example.com",
        to: str = "",
        sent: bool = False,
        received: dt.datetime | None = None,
        categories: str = "",
        unread: bool = False,
        importance: int = 1,
    ):
        self.EntryID = entry_id
        self.Subject = subject
        self.Body = body
        self.HTMLBody = body
        self.BodyFormat = None
        self.To = to
        self.CC = ""
        self.BCC = ""
        self.Categories = categories
        self.Importance = importance
        self.ConversationID = conversation_id
        self.SenderName = sender_name
        self.SenderEmailAddress = sender_address
        self.UnRead = unread
        self.FlagStatus = 0
        self.Attachments = FakeAttachments()
        self.Parent = folder
        self.Class = OL_CLASS_MAIL
        self.Saved = not sent
        self.Sent = sent
        self._timestamp = received
        self.ReceivedTime = received
        self.SentOn = received if sent else None
        self.LastModificationTime = received
        self.CreationTime = received
        self._sent_calls = 0
        folder.Items.append(self)

    def Save(self) -> None:
        self.Saved = True

    def Send(self) -> None:
        self.Sent = True
        self.Saved = False
        self._sent_calls += 1

    def Reply(self):
        return FakeMailItem(
            entry_id=f"{self.EntryID}-reply",
            subject=f"RE: {self.Subject}",
            body=f"Quoted: {self.Body}",
            folder=self.Parent,
            conversation_id=self.ConversationID,
            to=self.SenderEmailAddress,
            received=self._timestamp,
        )

    def ReplyAll(self):
        return self.Reply()

    def Forward(self):
        return FakeMailItem(
            entry_id=f"{self.EntryID}-fwd",
            subject=f"FW: {self.Subject}",
            body=f"Forwarded: {self.Body}",
            folder=self.Parent,
            conversation_id=self.ConversationID,
            received=self._timestamp,
        )


class FakeOutlook:
    def __init__(self, drafts_folder: FakeFolder, register):
        self._drafts_folder = drafts_folder
        self._register = register
        self._counter = 0

    def CreateItem(self, _item_type: int):
        self._counter += 1
        item = FakeMailItem(
            entry_id=f"draft-{self._counter}",
            subject="",
            body="",
            folder=self._drafts_folder,
            conversation_id=f"draft-conv-{self._counter}",
            received=dt.datetime(2026, 7, 25, 12, 0, self._counter),
        )
        self._register(item)
        return item


def _build_namespace():
    root = FakeFolder("Mailbox - test@example.com")
    inbox = FakeFolder("Inbox", root)
    sent = FakeFolder("Sent Items", root)
    drafts = FakeFolder("Drafts", root)
    deleted = FakeFolder("Deleted Items", root)

    items_by_id: dict[str, FakeMailItem] = {}

    def register(item: FakeMailItem) -> FakeMailItem:
        items_by_id[item.EntryID] = item
        return item

    register(
        FakeMailItem(
            "mail-1",
            "Budget thread",
            "Need the latest budget",
            inbox,
            conversation_id="conv-budget",
            sender_name="Alice",
            sender_address="alice@example.com",
            received=dt.datetime(2026, 7, 24, 9, 0, 0),
            categories="Finance, Urgent",
            unread=True,
        )
    )
    register(
        FakeMailItem(
            "mail-2",
            "RE: Budget thread",
            "Sent reply",
            sent,
            conversation_id="conv-budget",
            sender_name="Me",
            sender_address="me@example.com",
            to="alice@example.com",
            sent=True,
            received=dt.datetime(2026, 7, 24, 10, 0, 0),
            categories="Finance",
        )
    )
    register(
        FakeMailItem(
            "mail-3",
            "Another note",
            "Something else",
            drafts,
            conversation_id="conv-other",
            received=dt.datetime(2026, 7, 25, 8, 0, 0),
        )
    )

    default_map = {3: deleted, 5: sent, 6: inbox, 16: drafts}
    outlook = FakeOutlook(drafts, register)
    namespace = SimpleNamespace(
        Stores=[SimpleNamespace(DisplayName=root.Name, GetRootFolder=lambda: root)],
        GetDefaultFolder=lambda folder_id: default_map[folder_id],
        GetItemFromID=lambda entry_id, store_id=None: items_by_id[entry_id],
    )
    return outlook, namespace, items_by_id, {"root": root, "inbox": inbox, "sent": sent, "drafts": drafts, "deleted": deleted}


def test_send_mail_requires_explicit_confirmation():
    outlook, namespace, _, _ = _build_namespace()

    with pytest.raises(OutlookError, match="confirm_send=true"):
        mail_client.send_mail(
            outlook,
            namespace,
            to=["bob@example.com"],
            subject="Hello",
            body="Hi",
        )


def test_create_and_update_draft():
    outlook, namespace, items_by_id, _ = _build_namespace()

    payload = mail_client.create_draft(
        outlook,
        namespace,
        to=["bob@example.com"],
        subject="Draft",
        body="Body",
        categories="Ops",
    )
    updated = mail_client.update_draft(
        outlook,
        namespace,
        entry_id=payload["entry_id"],
        subject="Updated draft",
        body="<b>HTML</b>",
        html=True,
        categories="Ops, Review",
    )

    assert updated["status"] == "updated"
    item = namespace.GetItemFromID(payload["entry_id"])
    assert item.Subject == "Updated draft"
    assert item.HTMLBody == "<b>HTML</b>"
    assert item.Categories == "Ops, Review"


def test_send_draft_requires_confirmation_then_sends():
    outlook, namespace, items_by_id, _ = _build_namespace()
    draft = mail_client.create_draft(
        outlook,
        namespace,
        to=["bob@example.com"],
        subject="Draft to send",
        body="Body",
    )
    draft_item = namespace.GetItemFromID(draft["entry_id"])

    with pytest.raises(OutlookError, match="confirm_send=true"):
        mail_client.send_draft(outlook, namespace, entry_id=draft["entry_id"])

    payload = mail_client.send_draft(
        outlook, namespace, entry_id=draft["entry_id"], confirm_send=True
    )
    assert payload["status"] == "sent"
    assert draft_item.Sent is True


def test_list_conversation_collects_thread_across_folders():
    outlook, namespace, _, _ = _build_namespace()

    payload = mail_client.list_conversation(
        outlook, namespace, conversation_id="conv-budget"
    )

    assert payload["count"] == 2
    assert [item["entry_id"] for item in payload["items"]] == ["mail-1", "mail-2"]


def test_search_mails_supports_multi_folder_filters():
    outlook, namespace, _, _ = _build_namespace()

    payload = mail_client.search_mails(
        outlook,
        namespace,
        query="budget",
        folders=["inbox", "sent"],
        categories_contains=["Finance"],
        unread_only=False,
    )

    assert payload["count"] == 2
    assert {item["entry_id"] for item in payload["items"]} == {"mail-1", "mail-2"}
