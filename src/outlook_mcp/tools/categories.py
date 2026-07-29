"""MCP tool wrappers for Outlook categories."""

from __future__ import annotations

from typing import Annotated

from pydantic import Field

from outlook_mcp.client import categories as cat_client
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_list_categories",
        annotations={
            "title": "List Outlook color categories",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_list_categories(
        response_format: Annotated[
            str, Field(description="'markdown' or 'json'.")
        ] = "markdown",
    ) -> str:
        """List the color categories configured in this Outlook profile."""
        data = await bridge.call(cat_client.list_categories)
        return format_response(data, response_format)

    @mcp.tool(
        name="outlook_create_category",
        annotations={
            "title": "Create an Outlook color category",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_create_category(
        name: Annotated[str, Field(min_length=1, description="Category display name.")],
        color: Annotated[
            int | None,
            Field(
                description="Optional Outlook category color enum value. Omit to let Outlook choose.",
                ge=0,
                le=25,
            ),
        ] = None,
    ) -> str:
        """Create a profile-wide Outlook category."""
        data = await bridge.call(cat_client.create_category, name=name, color=color)
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_update_category",
        annotations={
            "title": "Rename or recolor an Outlook category",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_update_category(
        name: Annotated[
            str,
            Field(
                min_length=1,
                description="Exact current category name (use outlook_list_categories).",
            ),
        ],
        new_name: Annotated[
            str | None,
            Field(
                min_length=1,
                description="Optional replacement category name.",
            ),
        ] = None,
        color: Annotated[
            int | None,
            Field(
                ge=0,
                le=25,
                description="Optional replacement Outlook category color enum value.",
            ),
        ] = None,
    ) -> str:
        """Rename and/or recolor an existing profile-wide Outlook category."""
        data = await bridge.call(
            cat_client.update_category,
            name=name,
            new_name=new_name,
            color=color,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_set_category",
        annotations={
            "title": "Set categories on a mail / event / task",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_set_category(
        entry_id: Annotated[str, Field(description="EntryID of the item.")],
        categories: Annotated[
            str,
            Field(
                description=(
                    "Comma-separated category names (e.g. 'Important' or "
                    "'Work, Follow-up'). Empty string clears all categories."
                ),
            ),
        ],
    ) -> str:
        """Replace the categories on an item (mail / event / task)."""
        data = await bridge.call(
            cat_client.set_category, entry_id=entry_id, categories=categories
        )
        return format_response(data, "json")
