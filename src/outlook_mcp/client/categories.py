"""Category COM operations.

Outlook categories are profile-wide (not per-store), so we read them
straight off ``namespace.Categories``.
"""

from __future__ import annotations

from typing import Any

from outlook_mcp.errors import OutlookError
from outlook_mcp.client.folders import get_item_by_id


def _find_category(categories: Any, name: str) -> Any:
    for index in range(categories.Count):
        category = categories.Item(index + 1)
        if category.Name == name:
            return category
    raise OutlookError(
        f"Category '{name}' not found. Use outlook_list_categories to see available categories."
    )


def _validate_color(color: int) -> None:
    if not 0 <= color <= 25:
        raise OutlookError(
            "Category color must be an Outlook OlCategoryColor value from 0 to 25."
        )


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
    name = name.strip()
    if not name:
        raise OutlookError("Category name cannot be empty.")
    for i in range(cats.Count):
        cat = cats.Item(i + 1)
        if cat.Name.casefold() == name.casefold():
            raise OutlookError(
                f"Category '{name}' already exists. Use outlook_list_categories to inspect it."
            )
    if color is None:
        created = cats.Add(name)
    else:
        _validate_color(color)
        created = cats.Add(name, color)
    return {
        "status": "created",
        "category": {"name": created.Name, "color": created.Color},
    }


def update_category(
    outlook: Any,
    namespace: Any,
    *,
    name: str,
    new_name: str | None = None,
    color: int | None = None,
) -> dict[str, Any]:
    categories = namespace.Categories
    category = _find_category(categories, name)
    previous_name = category.Name

    if new_name is None and color is None:
        raise OutlookError("Provide new_name and/or color to update the category.")

    if new_name is not None:
        new_name = new_name.strip()
        if not new_name:
            raise OutlookError("new_name cannot be empty.")
        for index in range(categories.Count):
            existing = categories.Item(index + 1)
            if (
                existing.Name.casefold() == new_name.casefold()
                and existing.Name != previous_name
            ):
                raise OutlookError(
                    f"Category '{new_name}' already exists. Choose another name."
                )

    if color is not None:
        _validate_color(color)
    if new_name is not None:
        category.Name = new_name
    if color is not None:
        category.Color = color

    return {
        "status": "updated",
        "previous_name": previous_name,
        "category": {"name": category.Name, "color": category.Color},
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
