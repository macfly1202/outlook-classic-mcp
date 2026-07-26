"""MCP tool wrappers for mail."""

from __future__ import annotations

from typing import Annotated, Optional

from mcp.types import CallToolResult
from pydantic import Field

from outlook_mcp.client import mail as mail_client
from outlook_mcp.ui import ui_meta, ui_result
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_list_mails",
        annotations={
            "title": "List Outlook mails",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
        meta=ui_meta("mail-list"),
        structured_output=False,
    )
    @safe_call
    async def outlook_list_mails(
        folder: Annotated[
            str,
            Field(
                description=(
                    "Folder name. Either a well-known name (inbox, sent, drafts, "
                    "deleted, junk) or a path like 'Inbox/Projects/Quinn'."
                ),
            ),
        ] = "inbox",
        limit: Annotated[int, Field(ge=1, le=100, description="Max items.")] = 25,
        offset: Annotated[int, Field(ge=0, description="Pagination offset.")] = 0,
        unread_only: Annotated[bool, Field(description="Return only unread.")] = False,
        since: Annotated[Optional[str], Field(description="ISO-8601 lower bound on ReceivedTime.")] = None,
        until: Annotated[Optional[str], Field(description="ISO-8601 upper bound on ReceivedTime.")] = None,
        from_address: Annotated[Optional[str], Field(description="Substring match on sender email.")] = None,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> CallToolResult:
        """List mail items from a folder, newest first."""
        data = await bridge.call(
            mail_client.list_mails,
            folder=folder,
            limit=limit,
            offset=offset,
            unread_only=unread_only,
            since=since,
            until=until,
            from_address=from_address,
        )
        return ui_result(format_response(data, response_format), data)

    @mcp.tool(
        name="outlook_search_mails",
        annotations={
            "title": "Search Outlook mails",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
        meta=ui_meta("mail-list"),
        structured_output=False,
    )
    @safe_call
    async def outlook_search_mails(
        query: Annotated[str, Field(min_length=1, description="Search keywords or DASL filter.")],
        folder: Annotated[str, Field(description="Folder to search in.")] = "inbox",
        folders: Annotated[
            Optional[list[str]],
            Field(description="Optional list of folders to search across instead of a single folder."),
        ] = None,
        scope: Annotated[
            str,
            Field(
                description=(
                    "Where to look: 'subject_body' (default), 'subject', 'from', "
                    "or 'dasl' to pass `query` as a raw DASL @SQL filter."
                ),
            ),
        ] = "subject_body",
        limit: Annotated[int, Field(ge=1, le=100)] = 25,
        unread_only: Annotated[bool, Field(description="Only return unread mail.")] = False,
        since: Annotated[Optional[str], Field(description="ISO-8601 lower bound on mail timestamp.")] = None,
        until: Annotated[Optional[str], Field(description="ISO-8601 upper bound on mail timestamp.")] = None,
        from_address: Annotated[Optional[str], Field(description="Substring match on sender name/address.")] = None,
        has_attachments: Annotated[
            Optional[bool],
            Field(description="Filter to mail with or without attachments."),
        ] = None,
        importance: Annotated[
            Optional[str],
            Field(description="Optional importance filter: 'low', 'normal', or 'high'."),
        ] = None,
        categories_contains: Annotated[
            Optional[list[str]],
            Field(description="Require all listed categories on the mail."),
        ] = None,
        conversation_id: Annotated[
            Optional[str],
            Field(description="Restrict results to a single Outlook conversation/thread."),
        ] = None,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> CallToolResult:
        """Search a mail folder by subject, body, or sender."""
        data = await bridge.call(
            mail_client.search_mails,
            query=query,
            folder=folder,
            folders=folders,
            limit=limit,
            scope=scope,
            unread_only=unread_only,
            since=since,
            until=until,
            from_address=from_address,
            has_attachments=has_attachments,
            importance=importance,
            categories_contains=categories_contains,
            conversation_id=conversation_id,
        )
        return ui_result(format_response(data, response_format), data)

    @mcp.tool(
        name="outlook_get_mail",
        annotations={
            "title": "Get full Outlook mail",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
        meta=ui_meta("mail-view"),
        structured_output=False,
    )
    @safe_call
    async def outlook_get_mail(
        entry_id: Annotated[str, Field(min_length=1, description="EntryID of the mail item.")],
        include_body: Annotated[bool, Field(description="Include the plain-text body.")] = True,
        include_html: Annotated[
            bool,
            Field(
                description=(
                    "Also include the raw HTML body. Off by default — it is "
                    "usually huge and rarely needed; the plain-text body "
                    "carries the same content."
                ),
            ),
        ] = False,
        max_body_chars: Annotated[
            int,
            Field(
                ge=0,
                description="Truncate the body beyond this many chars (0 = no limit).",
            ),
        ] = 10000,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> CallToolResult:
        """Fetch body, headers, and attachment list for one mail item.

        If the response has body_truncated=true, re-call with a higher
        max_body_chars to read more.
        """
        data = await bridge.call(
            mail_client.get_mail,
            entry_id=entry_id,
            include_body=include_body,
            include_html=include_html,
            max_body_chars=max_body_chars,
        )
        return ui_result(format_response(data, response_format), data)

    @mcp.tool(
        name="outlook_send_mail",
        annotations={
            "title": "Send Outlook mail",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_send_mail(
        to: Annotated[list[str], Field(min_length=1, description="Recipient addresses.")],
        subject: Annotated[str, Field(description="Subject line.")],
        body: Annotated[str, Field(description="Message body. Plain text unless html=True.")],
        cc: Annotated[Optional[list[str]], Field(description="CC recipients.")] = None,
        bcc: Annotated[Optional[list[str]], Field(description="BCC recipients.")] = None,
        html: Annotated[bool, Field(description="Treat body as HTML.")] = False,
        attachments: Annotated[Optional[list[str]], Field(description="Absolute paths to local files.")] = None,
        importance: Annotated[str, Field(description="One of: 'low', 'normal', 'high'.")] = "normal",
        save_only: Annotated[bool, Field(description="If true, save to Drafts instead of sending.")] = False,
        confirm_send: Annotated[
            bool,
            Field(description="Required to actually send. Leave false to block accidental outbound mail."),
        ] = False,
    ) -> str:
        """Compose and send a new mail. Set save_only=True to save to Drafts."""
        data = await bridge.call(
            mail_client.send_mail,
            to=to,
            subject=subject,
            body=body,
            cc=cc,
            bcc=bcc,
            html=html,
            attachments=attachments,
            importance=importance,
            save_only=save_only,
            confirm_send=confirm_send,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_reply_mail",
        annotations={
            "title": "Reply to Outlook mail",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_reply_mail(
        entry_id: Annotated[str, Field(description="EntryID of the mail to reply to.")],
        body: Annotated[str, Field(description="Reply body. Quoted original is appended.")],
        reply_all: Annotated[bool, Field(description="Reply to all recipients.")] = False,
        html: Annotated[bool, Field(description="Treat body as HTML.")] = False,
        attachments: Annotated[Optional[list[str]], Field(description="Files to attach.")] = None,
        save_only: Annotated[bool, Field(description="Save the reply to Drafts instead of sending.")] = False,
        confirm_send: Annotated[
            bool,
            Field(description="Required to actually send the reply."),
        ] = False,
    ) -> str:
        """Reply (or reply-all) to an existing mail."""
        data = await bridge.call(
            mail_client.reply_mail,
            entry_id=entry_id,
            body=body,
            reply_all=reply_all,
            html=html,
            attachments=attachments,
            save_only=save_only,
            confirm_send=confirm_send,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_forward_mail",
        annotations={
            "title": "Forward Outlook mail",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_forward_mail(
        entry_id: Annotated[str, Field(description="EntryID of the mail to forward.")],
        to: Annotated[list[str], Field(min_length=1, description="Forward recipients.")],
        body: Annotated[str, Field(description="Optional note above the forwarded mail.")] = "",
        cc: Annotated[Optional[list[str]], Field(description="CC recipients.")] = None,
        html: Annotated[bool, Field(description="Treat body as HTML.")] = False,
        save_only: Annotated[bool, Field(description="Save the forward to Drafts instead of sending.")] = False,
        confirm_send: Annotated[
            bool,
            Field(description="Required to actually send the forward."),
        ] = False,
    ) -> str:
        """Forward an existing mail with an optional added note."""
        data = await bridge.call(
            mail_client.forward_mail,
            entry_id=entry_id,
            to=to,
            body=body,
            cc=cc,
            html=html,
            save_only=save_only,
            confirm_send=confirm_send,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_create_draft",
        annotations={
            "title": "Create Outlook draft",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_create_draft(
        to: Annotated[Optional[list[str]], Field(description="Optional recipient addresses.")] = None,
        subject: Annotated[str, Field(description="Draft subject line.")] = "",
        body: Annotated[str, Field(description="Draft body. Plain text unless html=True.")] = "",
        cc: Annotated[Optional[list[str]], Field(description="Optional CC recipients.")] = None,
        bcc: Annotated[Optional[list[str]], Field(description="Optional BCC recipients.")] = None,
        html: Annotated[bool, Field(description="Treat body as HTML.")] = False,
        attachments: Annotated[Optional[list[str]], Field(description="Absolute paths to local files.")] = None,
        importance: Annotated[str, Field(description="One of: 'low', 'normal', 'high'.")] = "normal",
        categories: Annotated[
            Optional[str],
            Field(description="Optional comma-separated categories to assign to the draft."),
        ] = None,
    ) -> str:
        """Create a draft in Outlook without sending it."""
        data = await bridge.call(
            mail_client.create_draft,
            to=to,
            subject=subject,
            body=body,
            cc=cc,
            bcc=bcc,
            html=html,
            attachments=attachments,
            importance=importance,
            categories=categories,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_update_draft",
        annotations={
            "title": "Update Outlook draft",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_update_draft(
        entry_id: Annotated[str, Field(description="EntryID of the draft mail item.")],
        to: Annotated[Optional[list[str]], Field(description="Optional replacement To list.")] = None,
        subject: Annotated[Optional[str], Field(description="Optional replacement subject line.")] = None,
        body: Annotated[Optional[str], Field(description="Optional replacement message body.")] = None,
        cc: Annotated[Optional[list[str]], Field(description="Optional replacement CC list.")] = None,
        bcc: Annotated[Optional[list[str]], Field(description="Optional replacement BCC list.")] = None,
        html: Annotated[
            Optional[bool],
            Field(description="Optional body format switch. If body is provided, controls plain text vs HTML."),
        ] = None,
        attachments_to_add: Annotated[
            Optional[list[str]],
            Field(description="Optional files to append to the draft."),
        ] = None,
        clear_attachments: Annotated[bool, Field(description="Remove all existing attachments first.")] = False,
        importance: Annotated[
            Optional[str],
            Field(description="Optional replacement importance: 'low', 'normal', or 'high'."),
        ] = None,
        categories: Annotated[
            Optional[str],
            Field(description="Optional replacement comma-separated categories."),
        ] = None,
    ) -> str:
        """Update an existing Outlook draft in place."""
        data = await bridge.call(
            mail_client.update_draft,
            entry_id=entry_id,
            to=to,
            subject=subject,
            body=body,
            cc=cc,
            bcc=bcc,
            html=html,
            attachments_to_add=attachments_to_add,
            clear_attachments=clear_attachments,
            importance=importance,
            categories=categories,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_send_draft",
        annotations={
            "title": "Send Outlook draft",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_send_draft(
        entry_id: Annotated[str, Field(description="EntryID of the draft mail item.")],
        confirm_send: Annotated[
            bool,
            Field(description="Required to actually send the draft."),
        ] = False,
    ) -> str:
        """Send a previously saved Outlook draft."""
        data = await bridge.call(
            mail_client.send_draft,
            entry_id=entry_id,
            confirm_send=confirm_send,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_list_conversation",
        annotations={
            "title": "List an Outlook mail conversation/thread",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
        meta=ui_meta("mail-list"),
        structured_output=False,
    )
    @safe_call
    async def outlook_list_conversation(
        entry_id: Annotated[
            Optional[str],
            Field(description="Seed mail EntryID. Use this or conversation_id."),
        ] = None,
        conversation_id: Annotated[
            Optional[str],
            Field(description="Explicit Outlook conversation/thread id."),
        ] = None,
        folders: Annotated[
            Optional[list[str]],
            Field(description="Optional folders to search across. Defaults to inbox, sent, drafts, deleted."),
        ] = None,
        limit: Annotated[int, Field(ge=1, le=200, description="Max items to return.")] = 100,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> CallToolResult:
        """List messages in one Outlook conversation/thread across common mail folders."""
        data = await bridge.call(
            mail_client.list_conversation,
            entry_id=entry_id,
            conversation_id=conversation_id,
            folders=folders,
            limit=limit,
        )
        return ui_result(format_response(data, response_format), data)

    @mcp.tool(
        name="outlook_move_mail",
        annotations={
            "title": "Move Outlook mail to folder",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_move_mail(
        entry_id: Annotated[str, Field(description="EntryID of the mail to move.")],
        target_folder: Annotated[str, Field(description="Destination folder.")],
    ) -> str:
        """Move a mail to another folder. Returns the new EntryID."""
        data = await bridge.call(
            mail_client.move_mail, entry_id=entry_id, target_folder=target_folder
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_delete_mail",
        annotations={
            "title": "Delete Outlook mail (move to Deleted Items)",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_delete_mail(
        entry_id: Annotated[str, Field(description="EntryID of the mail to delete.")],
    ) -> str:
        """Delete a mail (Outlook moves it to Deleted Items)."""
        data = await bridge.call(mail_client.delete_mail, entry_id=entry_id)
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_mark_mail",
        annotations={
            "title": "Mark Outlook mail read/unread or flag it",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_mark_mail(
        entry_id: Annotated[str, Field(description="EntryID of the mail.")],
        read: Annotated[Optional[bool], Field(description="True=mark read, False=unread, None=no change.")] = None,
        flagged: Annotated[Optional[bool], Field(description="True=flag for follow-up, False=clear flag.")] = None,
    ) -> str:
        """Toggle read state and/or follow-up flag on a mail."""
        data = await bridge.call(
            mail_client.mark_mail, entry_id=entry_id, read=read, flagged=flagged
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_save_attachments",
        annotations={
            "title": "Save Outlook mail attachments to disk",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_save_attachments(
        entry_id: Annotated[str, Field(description="EntryID of the mail.")],
        output_dir: Annotated[str, Field(description="Absolute directory under your user profile.")],
        attachment_index: Annotated[
            Optional[int], Field(ge=1, description="1-indexed attachment. Omit to save all.")
        ] = None,
    ) -> str:
        """Save one or all attachments from a mail to a local directory."""
        data = await bridge.call(
            mail_client.save_attachments,
            entry_id=entry_id,
            output_dir=output_dir,
            attachment_index=attachment_index,
        )
        return format_response(data, "json")
