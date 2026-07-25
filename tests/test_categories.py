"""Tests for category listing and creation without Outlook."""

from __future__ import annotations

import sys
import types
from types import SimpleNamespace

import pytest

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

    payload = categories_client.create_category(
        None, namespace, name="VIP", color=4
    )

    assert payload == {
        "status": "created",
        "category": {"name": "VIP", "color": 4},
    }


def test_create_category_rejects_duplicate_name():
    namespace = _namespace_with_categories([FakeCategory("VIP", 4)])

    with pytest.raises(OutlookError, match="already exists"):
        categories_client.create_category(None, namespace, name="VIP")
