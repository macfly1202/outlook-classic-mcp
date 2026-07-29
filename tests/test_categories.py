"""Tests for category management without Outlook."""

from __future__ import annotations

import sys
import types
from types import SimpleNamespace

import pytest

if sys.platform != "win32":
    sys.modules.setdefault("pythoncom", types.ModuleType("pythoncom"))

from outlook_mcp.client import categories as categories_client
from outlook_mcp.errors import OutlookError


class FakeCategory:
    def __init__(self, name: str, color: int):
        self.Name = name
        self.Color = color


class FakeCategories:
    def __init__(self, items: list[FakeCategory] | None = None):
        self._items = list(items or [])

    @property
    def Count(self) -> int:
        return len(self._items)

    def Item(self, index: int) -> FakeCategory:
        return self._items[index - 1]

    def Add(self, name: str, color: int | None = None) -> FakeCategory:
        created = FakeCategory(name, 25 if color is None else color)
        self._items.append(created)
        return created


def _namespace_with_categories(items: list[FakeCategory] | None = None):
    return SimpleNamespace(Categories=FakeCategories(items))


def test_list_categories_returns_profile_categories():
    namespace = _namespace_with_categories(
        [FakeCategory("Work", 1), FakeCategory("Urgent", 6)]
    )

    payload = categories_client.list_categories(None, namespace)

    assert payload["count"] == 2
    assert payload["items"][0] == {"name": "Work", "color": 1}


def test_create_category_uses_requested_color():
    namespace = _namespace_with_categories()

    payload = categories_client.create_category(None, namespace, name="VIP", color=4)

    assert payload == {
        "status": "created",
        "category": {"name": "VIP", "color": 4},
    }


def test_create_category_rejects_duplicate_name():
    namespace = _namespace_with_categories([FakeCategory("VIP", 4)])

    with pytest.raises(OutlookError, match="already exists"):
        categories_client.create_category(None, namespace, name="VIP")


def test_update_category_can_rename_and_recolor():
    namespace = _namespace_with_categories([FakeCategory("Old name", 4)])

    payload = categories_client.update_category(
        None,
        namespace,
        name="Old name",
        new_name="New name",
        color=8,
    )

    assert payload == {
        "status": "updated",
        "previous_name": "Old name",
        "category": {"name": "New name", "color": 8},
    }
    assert namespace.Categories.Item(1).Name == "New name"
    assert namespace.Categories.Item(1).Color == 8


def test_update_category_can_change_only_color():
    namespace = _namespace_with_categories([FakeCategory("Work", 4)])

    payload = categories_client.update_category(
        None,
        namespace,
        name="Work",
        color=16,
    )

    assert payload["previous_name"] == "Work"
    assert payload["category"] == {"name": "Work", "color": 16}


def test_update_category_rejects_duplicate_new_name():
    namespace = _namespace_with_categories(
        [FakeCategory("Work", 4), FakeCategory("Personal", 5)]
    )

    with pytest.raises(OutlookError, match="already exists"):
        categories_client.update_category(
            None,
            namespace,
            name="Work",
            new_name="Personal",
        )


def test_update_category_requires_a_change():
    namespace = _namespace_with_categories([FakeCategory("Work", 4)])

    with pytest.raises(OutlookError, match="new_name and/or color"):
        categories_client.update_category(None, namespace, name="Work")


def test_update_category_rejects_unknown_category():
    namespace = _namespace_with_categories()

    with pytest.raises(OutlookError, match="not found"):
        categories_client.update_category(
            None,
            namespace,
            name="Missing",
            new_name="Renamed",
        )


def test_category_color_must_be_valid_outlook_enum():
    namespace = _namespace_with_categories([FakeCategory("Work", 4)])

    with pytest.raises(OutlookError, match="0 to 25"):
        categories_client.update_category(
            None,
            namespace,
            name="Work",
            color=26,
        )
