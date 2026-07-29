"""MCP tool wrappers for mail rules."""

from __future__ import annotations

from typing import Annotated, Optional

from pydantic import Field

from outlook_mcp.client import rules as rules_client
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_list_rules",
        annotations={
            "title": "List Outlook mail rules",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_list_rules(
        response_format: Annotated[
            str, Field(description="'markdown' or 'json'.")
        ] = "markdown",
    ) -> str:
        """List all mail rules with status and supported editable fields."""
        data = await bridge.call(rules_client.list_rules)
        return format_response(data, response_format)

    @mcp.tool(
        name="outlook_toggle_rule",
        annotations={
            "title": "Enable or disable a mail rule",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_toggle_rule(
        rule_name: Annotated[
            str,
            Field(
                min_length=1, description="Exact rule name (use outlook_list_rules)."
            ),
        ],
        enabled: Annotated[
            bool, Field(description="True to enable, False to disable.")
        ],
    ) -> str:
        """Enable or disable a mail rule by name. Modifies live mail rules."""
        data = await bridge.call(
            rules_client.toggle_rule, rule_name=rule_name, enabled=enabled
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_create_rule",
        annotations={
            "title": "Create an Outlook mail rule",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_create_rule(
        name: Annotated[
            str, Field(min_length=1, description="Display name for the new rule.")
        ],
        sender_address_contains: Annotated[
            Optional[list[str]],
            Field(
                description="Optional sender-address substrings. Match is OR across values."
            ),
        ] = None,
        sent_to_recipients: Annotated[
            Optional[list[str]],
            Field(
                description="Optional recipients in To or Cc. Accepts resolvable names, aliases, or SMTP addresses; match is OR."
            ),
        ] = None,
        to_me: Annotated[
            Optional[bool],
            Field(description="Require the current mailbox to appear in the To field."),
        ] = None,
        subject_contains: Annotated[
            Optional[list[str]],
            Field(
                description="Optional subject substrings. Match is OR across values."
            ),
        ] = None,
        body_contains: Annotated[
            Optional[list[str]],
            Field(description="Optional body substrings. Match is OR across values."),
        ] = None,
        move_to_folder: Annotated[
            Optional[str],
            Field(description="Optional folder to move matching mail into."),
        ] = None,
        copy_to_folder: Annotated[
            Optional[str],
            Field(description="Optional folder to copy matching mail into."),
        ] = None,
        assign_categories: Annotated[
            Optional[list[str]],
            Field(description="Optional category names to assign to matching mail."),
        ] = None,
        except_sender_address_contains: Annotated[
            Optional[list[str]],
            Field(
                description="Optional sender substrings that prevent the rule from firing."
            ),
        ] = None,
        except_sent_to_recipients: Annotated[
            Optional[list[str]],
            Field(
                description="Recipients in To or Cc that prevent the rule from firing."
            ),
        ] = None,
        except_to_me: Annotated[
            Optional[bool],
            Field(
                description="Prevent the rule from firing when the current mailbox is in To."
            ),
        ] = None,
        except_subject_contains: Annotated[
            Optional[list[str]],
            Field(
                description="Optional subject substrings that prevent the rule from firing."
            ),
        ] = None,
        except_body_contains: Annotated[
            Optional[list[str]],
            Field(
                description="Optional body substrings that prevent the rule from firing."
            ),
        ] = None,
        stop_processing_more_rules: Annotated[
            bool,
            Field(
                description="Stop Outlook from evaluating later rules after this one matches."
            ),
        ] = False,
        execution_order: Annotated[
            Optional[int],
            Field(ge=1, description="Optional execution order among all rules."),
        ] = None,
        enabled: Annotated[
            bool, Field(description="Whether the new rule starts enabled.")
        ] = True,
    ) -> str:
        """Create a receive rule using supported COM-editable conditions and actions."""
        data = await bridge.call(
            rules_client.create_rule,
            name=name,
            sender_address_contains=sender_address_contains,
            sent_to_recipients=sent_to_recipients,
            to_me=to_me,
            subject_contains=subject_contains,
            body_contains=body_contains,
            move_to_folder=move_to_folder,
            copy_to_folder=copy_to_folder,
            assign_categories=assign_categories,
            except_sender_address_contains=except_sender_address_contains,
            except_sent_to_recipients=except_sent_to_recipients,
            except_to_me=except_to_me,
            except_subject_contains=except_subject_contains,
            except_body_contains=except_body_contains,
            stop_processing_more_rules=stop_processing_more_rules,
            execution_order=execution_order,
            enabled=enabled,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_update_rule",
        annotations={
            "title": "Update an Outlook mail rule",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_update_rule(
        rule_name: Annotated[
            str,
            Field(
                min_length=1,
                description="Exact current rule name (use outlook_list_rules).",
            ),
        ],
        new_name: Annotated[
            Optional[str],
            Field(description="Optional replacement display name."),
        ] = None,
        enabled: Annotated[
            Optional[bool],
            Field(description="Optional on/off change. Omit to leave unchanged."),
        ] = None,
        sender_address_contains: Annotated[
            Optional[list[str]],
            Field(
                description="Replace sender-address substrings. Pass [] to clear this condition."
            ),
        ] = None,
        sent_to_recipients: Annotated[
            Optional[list[str]],
            Field(
                description="Replace recipients matched in To or Cc; pass [] to clear."
            ),
        ] = None,
        to_me: Annotated[
            Optional[bool],
            Field(description="Enable or disable the current-mailbox-in-To condition."),
        ] = None,
        subject_contains: Annotated[
            Optional[list[str]],
            Field(
                description="Replace subject substrings. Pass [] to clear this condition."
            ),
        ] = None,
        body_contains: Annotated[
            Optional[list[str]],
            Field(
                description="Replace body substrings. Pass [] to clear this condition."
            ),
        ] = None,
        move_to_folder: Annotated[
            Optional[str],
            Field(description="Replace move target folder. Omit to leave unchanged."),
        ] = None,
        copy_to_folder: Annotated[
            Optional[str],
            Field(description="Replace copy target folder. Omit to leave unchanged."),
        ] = None,
        assign_categories: Annotated[
            Optional[list[str]],
            Field(
                description="Replace assigned categories. Pass [] to clear this action."
            ),
        ] = None,
        except_sender_address_contains: Annotated[
            Optional[list[str]],
            Field(description="Replace sender exceptions. Pass [] to clear."),
        ] = None,
        except_sent_to_recipients: Annotated[
            Optional[list[str]],
            Field(description="Replace recipient exceptions; pass [] to clear."),
        ] = None,
        except_to_me: Annotated[
            Optional[bool],
            Field(description="Enable or disable the current-mailbox-in-To exception."),
        ] = None,
        except_subject_contains: Annotated[
            Optional[list[str]],
            Field(description="Replace subject exceptions. Pass [] to clear."),
        ] = None,
        except_body_contains: Annotated[
            Optional[list[str]],
            Field(description="Replace body exceptions. Pass [] to clear."),
        ] = None,
        clear_move_to_folder: Annotated[
            bool,
            Field(description="Disable the move action entirely."),
        ] = False,
        clear_copy_to_folder: Annotated[
            bool,
            Field(description="Disable the copy action entirely."),
        ] = False,
        clear_assign_categories: Annotated[
            bool,
            Field(description="Disable category assignment entirely."),
        ] = False,
        stop_processing_more_rules: Annotated[
            Optional[bool],
            Field(description="Enable or disable stop-processing action."),
        ] = None,
        execution_order: Annotated[
            Optional[int],
            Field(ge=1, description="Optional replacement execution order."),
        ] = None,
    ) -> str:
        """Update a rule's supported fields: name, enabled state, conditions, and actions."""
        data = await bridge.call(
            rules_client.update_rule,
            rule_name=rule_name,
            new_name=new_name,
            enabled=enabled,
            sender_address_contains=sender_address_contains,
            sent_to_recipients=sent_to_recipients,
            to_me=to_me,
            subject_contains=subject_contains,
            body_contains=body_contains,
            move_to_folder=move_to_folder,
            copy_to_folder=copy_to_folder,
            assign_categories=assign_categories,
            except_sender_address_contains=except_sender_address_contains,
            except_sent_to_recipients=except_sent_to_recipients,
            except_to_me=except_to_me,
            except_subject_contains=except_subject_contains,
            except_body_contains=except_body_contains,
            clear_move_to_folder=clear_move_to_folder,
            clear_copy_to_folder=clear_copy_to_folder,
            clear_assign_categories=clear_assign_categories,
            stop_processing_more_rules=stop_processing_more_rules,
            execution_order=execution_order,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_delete_rule",
        annotations={
            "title": "Delete an Outlook mail rule",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_delete_rule(
        rule_name: Annotated[
            str,
            Field(
                min_length=1,
                description="Exact current rule name (use outlook_list_rules).",
            ),
        ],
    ) -> str:
        """Delete a live Outlook mail rule by name."""
        data = await bridge.call(
            rules_client.delete_rule,
            rule_name=rule_name,
        )
        return format_response(data, "json")
