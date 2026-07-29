"""Mail-rule COM operations.

Rules live on the default store. ``GetRules()`` returns a Rules
collection; ``Save()`` on the collection persists changes.

Programmatic rule creation in Outlook has only partial parity with the
Rules UI. This module intentionally exposes the most reliable subset:
receive rules with sender-address / recipient / subject / body conditions
and move/copy/category actions.
"""

from __future__ import annotations

from typing import Any

from outlook_mcp.client.folders import _safe_get, resolve_folder
from outlook_mcp.constants import OL_RULE_RECEIVE, OL_RULE_SEND
from outlook_mcp.errors import OutlookError


def _iter_rules(rules: Any):
    for i in range(rules.Count):
        yield i + 1, rules.Item(i + 1)


def _get_rules(namespace: Any) -> Any:
    return namespace.DefaultStore.GetRules()


def _find_rule(rules: Any, rule_name: str) -> tuple[int, Any]:
    for index, rule in _iter_rules(rules):
        if rule.Name == rule_name:
            return index, rule
    raise OutlookError(
        f"Rule '{rule_name}' not found. Use outlook_list_rules to see available rules."
    )


def _as_string_list(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value] if value.strip() else []
    if isinstance(value, (list, tuple)):
        return [str(item).strip() for item in value if str(item).strip()]
    try:
        return [str(item).strip() for item in value if str(item).strip()]
    except TypeError:
        text = str(value).strip()
        return [text] if text else []


def _clean_terms(values: list[str] | None) -> list[str]:
    if not values:
        return []
    return [value.strip() for value in values if value and value.strip()]


def _folder_label(folder: Any) -> str | None:
    if folder is None:
        return None
    path = _safe_get(folder, "FolderPath")
    if path:
        return str(path).lstrip("\\")
    return _safe_get(folder, "Name")


def _set_text_condition(condition: Any, values: list[str]) -> None:
    if values:
        condition.Text = list(values)
        condition.Enabled = True
    else:
        condition.Enabled = False
        try:
            condition.Text = []
        except Exception:
            pass


def _set_address_condition(condition: Any, values: list[str]) -> None:
    if values:
        condition.Address = list(values)
        condition.Enabled = True
    else:
        condition.Enabled = False
        try:
            condition.Address = []
        except Exception:
            pass


def _recipient_values(condition: Any) -> list[str]:
    recipients = _safe_get(condition, "Recipients")
    if recipients is None:
        return []
    values = []
    for index in range(1, int(_safe_get(recipients, "Count") or 0) + 1):
        recipient = recipients.Item(index)
        value = _safe_get(recipient, "Address") or _safe_get(recipient, "Name")
        if value:
            values.append(str(value))
    return values


def _clear_recipients(recipients: Any) -> None:
    while int(_safe_get(recipients, "Count") or 0):
        recipients.Remove(1)


def _set_recipient_condition(condition: Any, values: list[str]) -> None:
    recipients = condition.Recipients
    condition.Enabled = False
    _clear_recipients(recipients)
    if not values:
        return

    for value in values:
        recipients.Add(value)
    if not recipients.ResolveAll():
        unresolved = []
        for index in range(1, int(recipients.Count) + 1):
            recipient = recipients.Item(index)
            if not bool(_safe_get(recipient, "Resolved")):
                unresolved.append(
                    str(
                        _safe_get(recipient, "Name")
                        or _safe_get(recipient, "Address")
                        or values[index - 1]
                    )
                )
        _clear_recipients(recipients)
        raise OutlookError(
            "Outlook could not resolve these rule recipients: "
            + ", ".join(unresolved or values)
            + ". Use a resolvable display name, alias, or full SMTP address."
        )
    condition.Enabled = True


def _set_folder_action(namespace: Any, action: Any, folder_spec: str | None) -> None:
    if folder_spec:
        action.Folder = resolve_folder(namespace, folder_spec)
        action.Enabled = True
    else:
        action.Enabled = False
        try:
            action.Folder = None
        except Exception:
            pass


def _apply_supported_config(
    namespace: Any,
    target: Any,
    *,
    sender_address_contains: list[str] | None = None,
    sent_to_recipients: list[str] | None = None,
    to_me: bool | None = None,
    subject_contains: list[str] | None = None,
    body_contains: list[str] | None = None,
    move_to_folder: str | None = None,
    copy_to_folder: str | None = None,
    assign_categories: list[str] | None = None,
    clear_move_to_folder: bool = False,
    clear_copy_to_folder: bool = False,
    clear_assign_categories: bool = False,
    stop_processing_more_rules: bool | None = None,
) -> None:
    if sender_address_contains is not None:
        _set_address_condition(
            target.SenderAddress, _clean_terms(sender_address_contains)
        )
    if sent_to_recipients is not None:
        _set_recipient_condition(target.SentTo, _clean_terms(sent_to_recipients))
    if to_me is not None:
        target.ToMe.Enabled = to_me
    if subject_contains is not None:
        _set_text_condition(target.Subject, _clean_terms(subject_contains))
    if body_contains is not None:
        _set_text_condition(target.Body, _clean_terms(body_contains))
    if hasattr(target, "MoveToFolder") and (
        move_to_folder is not None or clear_move_to_folder
    ):
        _set_folder_action(
            namespace,
            target.MoveToFolder,
            None if clear_move_to_folder else move_to_folder,
        )
    if hasattr(target, "CopyToFolder") and (
        copy_to_folder is not None or clear_copy_to_folder
    ):
        _set_folder_action(
            namespace,
            target.CopyToFolder,
            None if clear_copy_to_folder else copy_to_folder,
        )
    if hasattr(target, "AssignToCategory") and (
        assign_categories is not None or clear_assign_categories
    ):
        categories = _clean_terms(assign_categories)
        action = target.AssignToCategory
        if clear_assign_categories or not categories:
            action.Enabled = False
            try:
                action.Categories = []
            except Exception:
                pass
        else:
            action.Categories = list(categories)
            action.Enabled = True
    if hasattr(target, "Stop") and stop_processing_more_rules is not None:
        target.Stop.Enabled = stop_processing_more_rules


def _rule_type_label(rule: Any) -> str:
    value = _safe_get(rule, "RuleType")
    if value == OL_RULE_RECEIVE:
        return "receive"
    if value == OL_RULE_SEND:
        return "send"
    return str(value)


def _rule_summary(index: int, rule: Any) -> dict[str, Any]:
    sender_condition = rule.Conditions.SenderAddress
    sent_to_condition = rule.Conditions.SentTo
    to_me_condition = rule.Conditions.ToMe
    subject_condition = rule.Conditions.Subject
    body_condition = rule.Conditions.Body
    move_action = rule.Actions.MoveToFolder
    copy_action = rule.Actions.CopyToFolder
    category_action = rule.Actions.AssignToCategory
    stop_action = rule.Actions.Stop
    return {
        "index": index,
        "name": rule.Name,
        "enabled": bool(rule.Enabled),
        "execution_order": _safe_get(rule, "ExecutionOrder"),
        "rule_type": _rule_type_label(rule),
        "supported_conditions": {
            "sender_address_contains": _as_string_list(
                _safe_get(sender_condition, "Address")
            )
            if _safe_get(sender_condition, "Enabled")
            else [],
            "sent_to_recipients": _recipient_values(sent_to_condition)
            if _safe_get(sent_to_condition, "Enabled")
            else [],
            "to_me": bool(_safe_get(to_me_condition, "Enabled")),
            "subject_contains": _as_string_list(_safe_get(subject_condition, "Text"))
            if _safe_get(subject_condition, "Enabled")
            else [],
            "body_contains": _as_string_list(_safe_get(body_condition, "Text"))
            if _safe_get(body_condition, "Enabled")
            else [],
        },
        "supported_exceptions": {
            "sender_address_contains": _as_string_list(
                _safe_get(rule.Exceptions.SenderAddress, "Address")
            )
            if _safe_get(rule.Exceptions.SenderAddress, "Enabled")
            else [],
            "sent_to_recipients": _recipient_values(rule.Exceptions.SentTo)
            if _safe_get(rule.Exceptions.SentTo, "Enabled")
            else [],
            "to_me": bool(_safe_get(rule.Exceptions.ToMe, "Enabled")),
            "subject_contains": _as_string_list(
                _safe_get(rule.Exceptions.Subject, "Text")
            )
            if _safe_get(rule.Exceptions.Subject, "Enabled")
            else [],
            "body_contains": _as_string_list(_safe_get(rule.Exceptions.Body, "Text"))
            if _safe_get(rule.Exceptions.Body, "Enabled")
            else [],
        },
        "supported_actions": {
            "move_to_folder": _folder_label(_safe_get(move_action, "Folder"))
            if _safe_get(move_action, "Enabled")
            else None,
            "copy_to_folder": _folder_label(_safe_get(copy_action, "Folder"))
            if _safe_get(copy_action, "Enabled")
            else None,
            "assign_categories": _as_string_list(
                _safe_get(category_action, "Categories")
            )
            if _safe_get(category_action, "Enabled")
            else [],
            "stop_processing_more_rules": bool(_safe_get(stop_action, "Enabled")),
        },
    }


def _validate_supported_rule_shape(rule: Any) -> None:
    supported_conditions = (
        bool(_safe_get(rule.Conditions.SenderAddress, "Enabled"))
        or bool(_safe_get(rule.Conditions.SentTo, "Enabled"))
        or bool(_safe_get(rule.Conditions.ToMe, "Enabled"))
        or bool(_safe_get(rule.Conditions.Subject, "Enabled"))
        or bool(_safe_get(rule.Conditions.Body, "Enabled"))
    )
    supported_actions = (
        bool(_safe_get(rule.Actions.MoveToFolder, "Enabled"))
        or bool(_safe_get(rule.Actions.CopyToFolder, "Enabled"))
        or bool(_safe_get(rule.Actions.AssignToCategory, "Enabled"))
    )
    if not supported_conditions:
        raise OutlookError(
            "At least one supported condition is required: sender_address_contains, "
            "sent_to_recipients, to_me, subject_contains, or body_contains."
        )
    if not supported_actions:
        raise OutlookError(
            "At least one supported action is required: move_to_folder or copy_to_folder."
            " or assign_categories."
        )


def list_rules(outlook: Any, namespace: Any) -> dict[str, Any]:
    rules = _get_rules(namespace)
    items = [_rule_summary(index, rule) for index, rule in _iter_rules(rules)]
    return {"count": len(items), "items": items}


def toggle_rule(
    outlook: Any,
    namespace: Any,
    *,
    rule_name: str,
    enabled: bool,
) -> dict[str, Any]:
    rules = _get_rules(namespace)
    index, rule = _find_rule(rules, rule_name)
    rule.Enabled = enabled
    rules.Save()
    return {
        "status": "updated",
        "rule": _rule_summary(index, rule),
    }


def create_rule(
    outlook: Any,
    namespace: Any,
    *,
    name: str,
    enabled: bool = True,
    sender_address_contains: list[str] | None = None,
    sent_to_recipients: list[str] | None = None,
    to_me: bool | None = None,
    subject_contains: list[str] | None = None,
    body_contains: list[str] | None = None,
    move_to_folder: str | None = None,
    copy_to_folder: str | None = None,
    assign_categories: list[str] | None = None,
    except_sender_address_contains: list[str] | None = None,
    except_sent_to_recipients: list[str] | None = None,
    except_to_me: bool | None = None,
    except_subject_contains: list[str] | None = None,
    except_body_contains: list[str] | None = None,
    stop_processing_more_rules: bool = False,
    execution_order: int | None = None,
) -> dict[str, Any]:
    rules = _get_rules(namespace)
    rule = rules.Create(name, OL_RULE_RECEIVE)
    _apply_supported_config(
        namespace,
        rule.Conditions,
        sender_address_contains=sender_address_contains,
        sent_to_recipients=sent_to_recipients,
        to_me=to_me,
        subject_contains=subject_contains,
        body_contains=body_contains,
    )
    _apply_supported_config(
        namespace,
        rule.Exceptions,
        sender_address_contains=except_sender_address_contains,
        sent_to_recipients=except_sent_to_recipients,
        to_me=except_to_me,
        subject_contains=except_subject_contains,
        body_contains=except_body_contains,
    )
    _apply_supported_config(
        namespace,
        rule.Actions,
        move_to_folder=move_to_folder,
        copy_to_folder=copy_to_folder,
        assign_categories=assign_categories,
        stop_processing_more_rules=stop_processing_more_rules,
    )
    _validate_supported_rule_shape(rule)
    if execution_order is not None:
        rule.ExecutionOrder = execution_order
    rule.Enabled = enabled
    rules.Save()
    index, saved_rule = _find_rule(rules, name)
    return {
        "status": "created",
        "rule": _rule_summary(index, saved_rule),
    }


def update_rule(
    outlook: Any,
    namespace: Any,
    *,
    rule_name: str,
    new_name: str | None = None,
    enabled: bool | None = None,
    sender_address_contains: list[str] | None = None,
    sent_to_recipients: list[str] | None = None,
    to_me: bool | None = None,
    subject_contains: list[str] | None = None,
    body_contains: list[str] | None = None,
    move_to_folder: str | None = None,
    copy_to_folder: str | None = None,
    assign_categories: list[str] | None = None,
    except_sender_address_contains: list[str] | None = None,
    except_sent_to_recipients: list[str] | None = None,
    except_to_me: bool | None = None,
    except_subject_contains: list[str] | None = None,
    except_body_contains: list[str] | None = None,
    clear_move_to_folder: bool = False,
    clear_copy_to_folder: bool = False,
    clear_assign_categories: bool = False,
    stop_processing_more_rules: bool | None = None,
    execution_order: int | None = None,
) -> dict[str, Any]:
    rules = _get_rules(namespace)
    _, rule = _find_rule(rules, rule_name)

    if new_name is not None:
        new_name = new_name.strip()
        if not new_name:
            raise OutlookError("new_name cannot be empty.")
        rule.Name = new_name
    if enabled is not None:
        rule.Enabled = enabled

    _apply_supported_config(
        namespace,
        rule.Conditions,
        sender_address_contains=sender_address_contains,
        sent_to_recipients=sent_to_recipients,
        to_me=to_me,
        subject_contains=subject_contains,
        body_contains=body_contains,
    )
    _apply_supported_config(
        namespace,
        rule.Exceptions,
        sender_address_contains=except_sender_address_contains,
        sent_to_recipients=except_sent_to_recipients,
        to_me=except_to_me,
        subject_contains=except_subject_contains,
        body_contains=except_body_contains,
    )
    _apply_supported_config(
        namespace,
        rule.Actions,
        move_to_folder=move_to_folder,
        copy_to_folder=copy_to_folder,
        assign_categories=assign_categories,
        clear_move_to_folder=clear_move_to_folder,
        clear_copy_to_folder=clear_copy_to_folder,
        clear_assign_categories=clear_assign_categories,
        stop_processing_more_rules=stop_processing_more_rules,
    )
    _validate_supported_rule_shape(rule)
    if execution_order is not None:
        rule.ExecutionOrder = execution_order
    rules.Save()
    final_name = new_name or rule_name
    index, saved_rule = _find_rule(rules, final_name)
    return {
        "status": "updated",
        "rule": _rule_summary(index, saved_rule),
    }


def delete_rule(
    outlook: Any,
    namespace: Any,
    *,
    rule_name: str,
) -> dict[str, Any]:
    rules = _get_rules(namespace)
    _, rule = _find_rule(rules, rule_name)
    removed_name = rule.Name
    rules.Remove(rule_name)
    rules.Save()
    return {
        "status": "deleted",
        "rule_name": removed_name,
    }
