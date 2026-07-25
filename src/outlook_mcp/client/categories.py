"""Category COM operations.

Outlook categories are profile-wide (not per-store), so we read them
straight off ``namespace.Categories``.
"""

from __future__ import annotations

from typing import Any

from outlook_mcp.errors import OutlookError
from outlook_mcp.client.folders import get_item_by_id


def list_categories(outlook: Any, namespace: Any) -> dict[str, Any]:
    items = []
    cats = namespace.Categories
    for i in range(cats.Count):
        cat = cats.Item(i + 1)
        items.append({"name": cat.Name, "color": cat.Color})
    return {"count": len(items), "items": items}


def create_category(
    outlook: Any,
    namespace: Any,
    *,
    name: str,
    color: int | None = None,
) -> dict[str, Any]:
    cats = namespace.Categories
    for i in range(cats.Count):
        cat = cats.Item(i + 1)
        if cat.Name == name:
            raise OutlookError(
                f"Category '{name}' already exists. Use outlook_list_categories to inspect it."
            )
    if color is None:
        created = cats.Add(name)
    else:
        created = cats.Add(name, color)
    return {
        "status": "created",
        "category": {"name": created.Name, "color": created.Color},
    }


def set_category(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    categories: str,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    item.Categories = categories
    item.Save()
    return {
        "status": "updated",
        "entry_id": entry_id,
        "subject": getattr(item, "Subject", ""),
        "categories": item.Categories or "",
    }
