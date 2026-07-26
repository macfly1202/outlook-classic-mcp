"""Tests for rule listing, creation, and updates without Outlook."""

from __future__ import annotations

import sys
import types
from types import SimpleNamespace

import pytest

sys.modules.setdefault("pythoncom", types.ModuleType("pythoncom"))

from outlook_mcp.client import rules as rules_client
from outlook_mcp.constants import OL_RULE_RECEIVE
from outlook_mcp.errors import OutlookError


class FakeFolder:
    def __init__(self, name: str, parent: "FakeFolder | None" = None):
        self.Name = name
        self.Parent = parent
        self.Folders: list[FakeFolder] = []
        self.Items = SimpleNamespace(Count=0)
        self.UnReadItemCount = 0
        self.DefaultItemType = 0
        if parent is None:
            self.FolderPath = f"\\\\{name}"
        else:
            self.FolderPath = f"{parent.FolderPath}\\{name}"
            parent.Folders.append(self)


class FakeTextCondition:
    def __init__(self, enabled: bool = False, text: list[str] | None = None):
        self.Enabled = enabled
        self.Text = list(text or [])


class FakeAddressCondition:
    def __init__(self, enabled: bool = False, address: list[str] | None = None):
        self.Enabled = enabled
        self.Address = list(address or [])


class FakeFolderAction:
    def __init__(self, enabled: bool = False, folder: FakeFolder | None = None):
        self.Enabled = enabled
        self.Folder = folder


class FakeAssignToCategoryAction:
    def __init__(self, enabled: bool = False, categories: list[str] | None = None):
        self.Enabled = enabled
        self.Categories = list(categories or [])


class FakeRuleAction:
    def __init__(self, enabled: bool = False):
        self.Enabled = enabled


class FakeConditions:
    def __init__(self):
        self.SenderAddress = FakeAddressCondition()
        self.Subject = FakeTextCondition()
        self.Body = FakeTextCondition()


class FakeActions:
    def __init__(self):
        self.MoveToFolder = FakeFolderAction()
        self.CopyToFolder = FakeFolderAction()
        self.AssignToCategory = FakeAssignToCategoryAction()
        self.Stop = FakeRuleAction()


class FakeRule:
    def __init__(self, name: str, rule_type: int = OL_RULE_RECEIVE):
        self.Name = name
        self.RuleType = rule_type
        self.Enabled = True
        self._execution_order = 1
        self.Conditions = FakeConditions()
        self.Exceptions = FakeConditions()
        self.Actions = FakeActions()

    @property
    def ExecutionOrder(self) -> int:
        return self._execution_order

    @ExecutionOrder.setter
    def ExecutionOrder(self, value: int) -> None:
        self._execution_order = value


class FakeRules:
    def __init__(self, items: list[FakeRule] | None = None):
        self._items = list(items or [])
        self.save_count = 0
        self._refresh_orders()

    @property
    def Count(self) -> int:
        return len(self._items)

    def Item(self, index: int) -> FakeRule:
        return self._items[index - 1]

    def Create(self, name: str, rule_type: int) -> FakeRule:
        rule = FakeRule(name, rule_type=rule_type)
        self._items.append(rule)
        self._refresh_orders()
        return rule

    def Save(self) -> None:
        self.save_count += 1
        desired_orders = {id(rule): rule.ExecutionOrder for rule in self._items}
        for rule in list(self._items):
            desired = desired_orders[id(rule)]
            current = self._items.index(rule) + 1
            if desired != current:
                self._items.remove(rule)
                insert_at = max(0, min(desired - 1, len(self._items)))
                self._items.insert(insert_at, rule)
        self._refresh_orders()

    def Remove(self, index: str | int) -> None:
        if isinstance(index, int):
            del self._items[index - 1]
            self._refresh_orders()
            return
        for i, rule in enumerate(self._items):
            if rule.Name == index:
                del self._items[i]
                self._refresh_orders()
                return
        raise KeyError(index)

    def _refresh_orders(self) -> None:
        for i, rule in enumerate(self._items, start=1):
            rule.ExecutionOrder = i


class FakeStore:
    def __init__(self, name: str, rules: FakeRules, root: FakeFolder):
        self.DisplayName = name
        self._rules = rules
        self._root = root

    def GetRules(self) -> FakeRules:
        return self._rules

    def GetRootFolder(self) -> FakeFolder:
        return self._root


def _namespace_with_rules(items: list[FakeRule] | None = None):
    root = FakeFolder("Mailbox - test@example.com")
    inbox = FakeFolder("Inbox", root)
    projects = FakeFolder("Projects", inbox)
    archive = FakeFolder("Archive", inbox)
    rules = FakeRules(items)
    store = FakeStore(root.Name, rules, root)
    namespace = SimpleNamespace(
        DefaultStore=store,
        Stores=[store],
        GetDefaultFolder=lambda folder_id: inbox,
    )
    return namespace, rules, {"root": root, "inbox": inbox, "projects": projects, "archive": archive}


def test_list_rules_includes_supported_details():
    rule = FakeRule("Route alerts")
    rule.Conditions.SenderAddress.Enabled = True
    rule.Conditions.SenderAddress.Address = ["alerts@example.com"]
    rule.Conditions.Subject.Enabled = True
    rule.Conditions.Subject.Text = ["urgent"]
    namespace, _, folders = _namespace_with_rules([rule])
    rule.Actions.MoveToFolder.Enabled = True
    rule.Actions.MoveToFolder.Folder = folders["projects"]

    payload = rules_client.list_rules(None, namespace)

    assert payload["count"] == 1
    item = payload["items"][0]
    assert item["name"] == "Route alerts"
    assert item["supported_conditions"]["sender_address_contains"] == [
        "alerts@example.com"
    ]
    assert item["supported_conditions"]["subject_contains"] == ["urgent"]
    assert item["supported_exceptions"]["subject_contains"] == []
    assert item["supported_actions"]["move_to_folder"] == (
        "Mailbox - test@example.com\\Inbox\\Projects"
    )


def test_create_rule_persists_conditions_and_actions():
    namespace, rules, _ = _namespace_with_rules()

    payload = rules_client.create_rule(
        None,
        namespace,
        name="File invoices",
        sender_address_contains=["billing@example.com"],
        subject_contains=["Invoice"],
        move_to_folder="Inbox/Projects",
    )

    assert payload["status"] == "created"
    assert rules.save_count == 1
    created = rules.Item(1)
    assert created.Name == "File invoices"
    assert created.Conditions.SenderAddress.Address == ["billing@example.com"]
    assert created.Conditions.Subject.Text == ["Invoice"]
    assert created.Actions.MoveToFolder.Folder.Name == "Projects"


def test_update_rule_can_rename_and_replace_supported_fields():
    rule = FakeRule("Jira")
    rule.Conditions.Subject.Enabled = True
    rule.Conditions.Subject.Text = ["Jira"]
    namespace, rules, folders = _namespace_with_rules([rule])
    rule.Actions.MoveToFolder.Enabled = True
    rule.Actions.MoveToFolder.Folder = folders["projects"]

    payload = rules_client.update_rule(
        None,
        namespace,
        rule_name="Jira",
        new_name="Jira triage",
        enabled=False,
        sender_address_contains=["jira@example.com"],
        subject_contains=[],
        except_subject_contains=["ignore"],
        copy_to_folder="Inbox/Archive",
        clear_move_to_folder=True,
        stop_processing_more_rules=True,
    )

    assert payload["status"] == "updated"
    assert rules.save_count == 1
    updated = rules.Item(1)
    assert updated.Name == "Jira triage"
    assert updated.Enabled is False
    assert updated.Conditions.SenderAddress.Address == ["jira@example.com"]
    assert updated.Conditions.Subject.Enabled is False
    assert updated.Exceptions.Subject.Text == ["ignore"]
    assert updated.Actions.MoveToFolder.Enabled is False
    assert updated.Actions.CopyToFolder.Enabled is True
    assert updated.Actions.CopyToFolder.Folder.Name == "Archive"
    assert updated.Actions.Stop.Enabled is True


def test_create_rule_can_assign_categories():
    namespace, rules, _ = _namespace_with_rules()

    payload = rules_client.create_rule(
        None,
        namespace,
        name="Tag VIP",
        sender_address_contains=["vip@example.com"],
        assign_categories=["VIP", "Follow-up"],
    )

    assert payload["status"] == "created"
    created = rules.Item(1)
    assert created.Actions.AssignToCategory.Enabled is True
    assert created.Actions.AssignToCategory.Categories == ["VIP", "Follow-up"]


def test_create_rule_can_set_order_and_exceptions():
    first = FakeRule("First")
    first.Conditions.Subject.Enabled = True
    first.Conditions.Subject.Text = ["a"]
    first.Actions.MoveToFolder.Enabled = True
    second = FakeRule("Second")
    second.Conditions.Subject.Enabled = True
    second.Conditions.Subject.Text = ["b"]
    namespace, rules, folders = _namespace_with_rules([first, second])
    first.Actions.MoveToFolder.Folder = folders["projects"]
    second.Actions.MoveToFolder.Enabled = True
    second.Actions.MoveToFolder.Folder = folders["archive"]

    payload = rules_client.create_rule(
        None,
        namespace,
        name="Inserted",
        sender_address_contains=["alerts@example.com"],
        except_body_contains=["skip"],
        move_to_folder="Inbox/Projects",
        stop_processing_more_rules=True,
        execution_order=2,
    )

    assert payload["status"] == "created"
    assert rules.Item(2).Name == "Inserted"
    assert rules.Item(2).Exceptions.Body.Text == ["skip"]
    assert rules.Item(2).Actions.Stop.Enabled is True


def test_delete_rule_removes_rule_by_name():
    first = FakeRule("First")
    second = FakeRule("Second")
    first.Conditions.Subject.Enabled = True
    first.Conditions.Subject.Text = ["a"]
    second.Conditions.Subject.Enabled = True
    second.Conditions.Subject.Text = ["b"]
    namespace, rules, folders = _namespace_with_rules([first, second])
    first.Actions.MoveToFolder.Enabled = True
    first.Actions.MoveToFolder.Folder = folders["projects"]
    second.Actions.MoveToFolder.Enabled = True
    second.Actions.MoveToFolder.Folder = folders["archive"]

    payload = rules_client.delete_rule(None, namespace, rule_name="First")

    assert payload == {"status": "deleted", "rule_name": "First"}
    assert rules.Count == 1
    assert rules.Item(1).Name == "Second"


def test_create_rule_requires_supported_action():
    namespace, _, _ = _namespace_with_rules()

    with pytest.raises(OutlookError, match="supported action"):
        rules_client.create_rule(
            None,
            namespace,
            name="Broken",
            sender_address_contains=["alerts@example.com"],
        )


def test_update_rule_rejects_removing_last_supported_condition():
    rule = FakeRule("Only move")
    rule.Conditions.Subject.Enabled = True
    rule.Conditions.Subject.Text = ["foo"]
    namespace, _, folders = _namespace_with_rules([rule])
    rule.Actions.MoveToFolder.Enabled = True
    rule.Actions.MoveToFolder.Folder = folders["projects"]

    with pytest.raises(OutlookError, match="supported condition"):
        rules_client.update_rule(
            None,
            namespace,
            rule_name="Only move",
            subject_contains=[],
        )
