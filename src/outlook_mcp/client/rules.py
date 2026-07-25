"""Mail-rule COM operations.

Rules live on the default store. ``GetRules()`` returns a Rules
collection; ``Save()`` on the collection persists changes.

Programmatic rule creation in Outlook has only partial parity with the
Rules UI. This module intentionally exposes the most reliable subset:
receive rules with sender-address / subject / body conditions and
move/copy actions.
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
    rule: Any,
    *,
    sender_address_contains: list[str] | None = None,
    subject_contains: list[str] | None = None,
    body_contains: list[str] | None = None,
    move_to_folder: str | None = None,
    copy_to_folder: str | None = None,
    assign_categories: list[str] | None = None,
    clear_move_to_folder: bool = False,
    clear_copy_to_folder: bool = False,
    clear_assign_categories: bool = False,
) -> None:
    if sender_address_contains is not None:
        _set_address_condition(
            rule.Conditions.SenderAddress, _clean_terms(sender_address_contains)
        )
    if subject_contains is not None:
        _set_text_condition(rule.Conditions.Subject, _clean_terms(subject_contains))
    if body_contains is not None:
        _set_text_condition(rule.Conditions.Body, _clean_terms(body_contains))
    if move_to_folder is not None or clear_move_to_folder:
        _set_folder_action(
            namespace,
            rule.Actions.MoveToFolder,
            None if clear_move_to_folder else move_to_folder,
        )
    if copy_to_folder is not None or clear_copy_to_folder:
        _set_folder_action(
            namespace,
            rule.Actions.CopyToFolder,
            None if clear_copy_to_folder else copy_to_folder,
        )
    if assign_categories is not None or clear_assign_categories:
        categories = _clean_terms(assign_categories)
        action = rule.Actions.AssignToCategory
        if clear_assign_categories or not categories:
            action.Enabled = False
            try:
                action.Categories = []
            except Exception:
                pass
        else:
            action.Categories = list(categories)
            action.Enabled = True


def _rule_type_label(rule: Any) -> str:
    value = _safe_get(rule, "RuleType")
    if value == OL_RULE_RECEIVE:
        return "receive"
    if value == OL_RULE_SEND:
        return "send"
    return str(value)


def _rule_summary(index: int, rule: Any) -> dict[str, Any]:
    sender_condition = rule.Conditions.SenderAddress
    subject_condition = rule.Conditions.Subject
    body_condition = rule.Conditions.Body
    move_action = rule.Actions.MoveToFolder
    copy_action = rule.Actions.CopyToFolder
    category_action = rule.Actions.AssignToCategory
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
            "subject_contains": _as_string_list(_safe_get(subject_condition, "Text"))
            if _safe_get(subject_condition, "Enabled")
            else [],
            "body_contains": _as_string_list(_safe_get(body_condition, "Text"))
            if _safe_get(body_condition, "Enabled")
            else [],
        },
        "supported_actions": {
            "move_to_folder": _folder_label(_safe_get(move_action, "Folder"))
            if _safe_get(move_action, "Enabled")
            else None,
            "copy_to_folder": _folder_label(_safe_get(copy_action, "Folder"))
            if _safe_get(copy_action, "Enabled")
            else None,
            "assign_categories": _as_string_list(_safe_get(category_action, "Categories"))
            if _safe_get(category_action, "Enabled")
            else [],
        },
    }


def _validate_supported_rule_shape(rule: Any) -> None:
    supported_conditions = (
        bool(_safe_get(rule.Conditions.SenderAddress, "Enabled"))
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
            "subject_contains, or body_contains."
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
    subject_contains: list[str] | None = None,
    body_contains: list[str] | None = None,
    move_to_folder: str | None = None,
    copy_to_folder: str | None = None,
    assign_categories: list[str] | None = None,
) -> dict[str, Any]:
    rules = _get_rules(namespace)
    rule = rules.Create(name, OL_RULE_RECEIVE)
    _apply_supported_config(
        namespace,
        rule,
        sender_address_contains=sender_address_contains,
        subject_contains=subject_contains,
        body_contains=body_contains,
        move_to_folder=move_to_folder,
        copy_to_folder=copy_to_folder,
        assign_categories=assign_categories,
    )
    _validate_supported_rule_shape(rule)
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
    subject_contains: list[str] | None = None,
    body_contains: list[str] | None = None,
    move_to_folder: str | None = None,
    copy_to_folder: str | None = None,
    assign_categories: list[str] | None = None,
    clear_move_to_folder: bool = False,
    clear_copy_to_folder: bool = False,
    clear_assign_categories: bool = False,
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
        rule,
        sender_address_contains=sender_address_contains,
        subject_contains=subject_contains,
        body_contains=body_contains,
        move_to_folder=move_to_folder,
        copy_to_folder=copy_to_folder,
        assign_categories=assign_categories,
        clear_move_to_folder=clear_move_to_folder,
        clear_copy_to_folder=clear_copy_to_folder,
        clear_assign_categories=clear_assign_categories,
    )
    _validate_supported_rule_shape(rule)
    rules.Save()
    final_name = new_name or rule_name
    index, saved_rule = _find_rule(rules, final_name)
    return {
        "status": "updated",
        "rule": _rule_summary(index, saved_rule),
    }
