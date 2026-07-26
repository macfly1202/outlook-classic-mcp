"""Mail COM operations."""

from __future__ import annotations

import datetime as dt
import ntpath
import os
from typing import Any

from outlook_mcp.client.folders import _safe_get, get_item_by_id, resolve_folder
from outlook_mcp.constants import (
    IMPORTANCE_MAP,
    OL_CLASS_MAIL,
    OL_CLASS_MEETING_REQUEST,
    OL_FORMAT_HTML,
    OL_FORMAT_PLAIN,
    OL_IMPORTANCE_NORMAL,
    OL_MAIL_ITEM,
)
from outlook_mcp.errors import OutlookError
from outlook_mcp.utils.formatting import from_iso, to_iso, truncate
from outlook_mcp.utils.paths import validate_attachment_path, validate_output_dir
from outlook_mcp.utils.safety import safe_dasl

WINDOWS_RESERVED_DEVICE_NAMES = {"CON", "PRN", "AUX", "NUL", "CLOCK$"} | {
    f"COM{i}" for i in range(1, 10)
} | {f"LPT{i}" for i in range(1, 10)}

# PR_SENDER_SMTP_ADDRESS — unlike urn:schemas:httpmail:fromemail, this
# holds the real SMTP address even for Exchange senders (whose fromemail
# is an EX:/O=... distinguished name).
SMTP_PROPTAG = "http://schemas.microsoft.com/mapi/proptag/0x5D02001F"
DEFAULT_CONVERSATION_FOLDERS = ["inbox", "sent", "drafts", "deleted"]


def split_search_words(query: str) -> tuple[str, list[str]]:
    """Split a search query into (anchor, remaining_words), lowercased.

    The anchor is the longest word — the most selective term to push
    down into the DASL Restrict. The remaining words are verified in
    Python per item, because DASL can't reliably AND two LIKEs on the
    same property (verified live: it returns zero rows).
    """
    words = [w.lower() for w in query.split() if w]
    if not words:
        return query.lower(), []
    anchor = max(words, key=len)
    remaining = list(words)
    remaining.remove(anchor)
    return anchor, remaining


def _search_haystack(item: Any, scope: str) -> str:
    if scope == "subject":
        fields = ("Subject",)
    elif scope == "from":
        fields = ("SenderName", "SenderEmailAddress")
    else:  # subject_body
        fields = ("Subject", "Body")
    return " ".join(str(_safe_get(item, f, "") or "") for f in fields).lower()


def _mail_timestamp(item: Any) -> dt.datetime | None:
    for attr in ("ReceivedTime", "SentOn", "LastModificationTime", "CreationTime"):
        value = _safe_get(item, attr)
        if isinstance(value, dt.datetime):
            if value.tzinfo is not None:
                value = value.astimezone().replace(tzinfo=None)
            return value
    return None


def _folder_path(folder: Any) -> str | None:
    path = _safe_get(folder, "FolderPath")
    if path:
        return str(path).lstrip("\\")
    return _safe_get(folder, "Name")


def _iter_collection(collection: Any):
    if collection is None:
        return iter(())
    try:
        return iter(collection)
    except TypeError:
        items = []
        count = _safe_get(collection, "Count", 0) or 0
        item_method = _safe_get(collection, "Item")
        if callable(item_method):
            for i in range(count):
                items.append(item_method(i + 1))
        return iter(items)


def _resolve_mail_folders(
    namespace: Any,
    *,
    folder: str | None = None,
    folders: list[str] | None = None,
) -> list[Any]:
    specs = folders or ([folder] if folder else [])
    if not specs:
        specs = ["inbox"]
    resolved = []
    seen: set[str] = set()
    for spec in specs:
        resolved_folder = resolve_folder(namespace, spec)
        key = _folder_path(resolved_folder) or str(id(resolved_folder))
        if key in seen:
            continue
        seen.add(key)
        resolved.append(resolved_folder)
    return resolved


def _mail_matches_filters(
    item: Any,
    *,
    query: str | None = None,
    scope: str = "subject_body",
    remaining_words: list[str] | None = None,
    unread_only: bool = False,
    since_dt: dt.datetime | None = None,
    until_dt: dt.datetime | None = None,
    from_address: str | None = None,
    has_attachments: bool | None = None,
    importance: str | None = None,
    categories_contains: list[str] | None = None,
    conversation_id: str | None = None,
) -> bool:
    cls = _safe_get(item, "Class")
    if cls not in (OL_CLASS_MAIL, OL_CLASS_MEETING_REQUEST):
        return False

    if unread_only and not bool(_safe_get(item, "UnRead", False)):
        return False

    item_ts = _mail_timestamp(item)
    if since_dt and item_ts and item_ts < since_dt:
        return False
    if until_dt and item_ts and item_ts > until_dt:
        return False

    if from_address:
        sender = " ".join(
            [
                str(_safe_get(item, "SenderEmailAddress", "") or ""),
                str(_safe_get(item, "SenderName", "") or ""),
            ]
        ).lower()
        if from_address.lower() not in sender:
            return False

    attachments = _safe_get(item, "Attachments")
    item_has_attachments = attachments.Count > 0 if attachments else False
    if has_attachments is not None and item_has_attachments != has_attachments:
        return False

    if importance:
        wanted = IMPORTANCE_MAP.get(importance.lower())
        if wanted is None or _safe_get(item, "Importance") != wanted:
            return False

    if categories_contains:
        item_categories = [
            part.strip().lower()
            for part in str(_safe_get(item, "Categories", "") or "").split(",")
            if part.strip()
        ]
        wanted_categories = [part.strip().lower() for part in categories_contains if part.strip()]
        if not all(category in item_categories for category in wanted_categories):
            return False

    if conversation_id and _safe_get(item, "ConversationID") != conversation_id:
        return False

    if query:
        haystack = _search_haystack(item, scope)
        anchor, fallback_remaining = split_search_words(query)
        words = [anchor] + list(remaining_words or fallback_remaining)
        if scope != "dasl" and not all(word in haystack for word in words if word):
            return False

    return True


def _require_send_confirmation(*, confirm_send: bool, action: str) -> None:
    if not confirm_send:
        raise OutlookError(
            f"{action} is blocked until you explicitly confirm the send. "
            "Re-run with confirm_send=true, or use save_only=true to stage a draft."
        )


def _mail_summary(item: Any) -> dict[str, Any]:
    attachments = _safe_get(item, "Attachments")
    return {
        "entry_id": _safe_get(item, "EntryID"),
        "conversation_id": _safe_get(item, "ConversationID"),
        "subject": _safe_get(item, "Subject", ""),
        "from": _safe_get(item, "SenderName"),
        "from_address": _safe_get(item, "SenderEmailAddress"),
        "to": _safe_get(item, "To", ""),
        "received": to_iso(_safe_get(item, "ReceivedTime")),
        "sent": to_iso(_safe_get(item, "SentOn")),
        "unread": bool(_safe_get(item, "UnRead", False)),
        "flagged": _safe_get(item, "FlagStatus") == 2,  # olFlagMarked
        "has_attachments": attachments.Count > 0 if attachments else False,
        "importance": _safe_get(item, "Importance"),
        "categories": _safe_get(item, "Categories", "") or "",
        "folder_path": _folder_path(_safe_get(item, "Parent")),
        "is_draft": bool(_safe_get(item, "Saved", False) and not _safe_get(item, "Sent", False)),
        "preview": truncate(_safe_get(item, "Body", ""), 200),
    }


def _mail_full(
    item: Any,
    include_body: bool = True,
    include_html: bool = False,
    max_body_chars: int = 10000,
) -> dict[str, Any]:
    attachments = []
    if _safe_get(item, "Attachments"):
        for i, att in enumerate(item.Attachments, start=1):
            attachments.append(
                {
                    "index": i,
                    "filename": att.FileName,
                    "size_bytes": _safe_get(att, "Size"),
                }
            )
    result = {
        "entry_id": _safe_get(item, "EntryID"),
        "conversation_id": _safe_get(item, "ConversationID"),
        "subject": _safe_get(item, "Subject", ""),
        "from": _safe_get(item, "SenderName"),
        "from_address": _safe_get(item, "SenderEmailAddress"),
        "to": _safe_get(item, "To", ""),
        "cc": _safe_get(item, "CC", ""),
        "bcc": _safe_get(item, "BCC", ""),
        "received": to_iso(_safe_get(item, "ReceivedTime")),
        "sent": to_iso(_safe_get(item, "SentOn")),
        "unread": bool(_safe_get(item, "UnRead", False)),
        "flagged": _safe_get(item, "FlagStatus") == 2,  # olFlagMarked
        "importance": _safe_get(item, "Importance"),
        "categories": _safe_get(item, "Categories", ""),
        "attachments": attachments,
    }
    if include_body:
        body = _safe_get(item, "Body", "") or ""
        if max_body_chars and len(body) > max_body_chars:
            result["body"] = body[:max_body_chars].rstrip()
            result["body_truncated"] = True
            result["body_total_chars"] = len(body)
        else:
            result["body"] = body
    if include_html:
        # Full HTML of a styled corporate mail easily runs to tens of
        # kilobytes — only fetch when explicitly asked for.
        result["html_body"] = _safe_get(item, "HTMLBody", "")
    return result


def list_mails(
    outlook: Any,
    namespace: Any,
    *,
    folder: str | None = "inbox",
    limit: int = 25,
    offset: int = 0,
    unread_only: bool = False,
    since: str | None = None,
    until: str | None = None,
    from_address: str | None = None,
) -> dict[str, Any]:
    f = resolve_folder(namespace, folder)
    items = f.Items
    items.Sort("[ReceivedTime]", True)

    clauses: list[str] = []
    if unread_only:
        clauses.append("[UnRead] = True")
    since_dt = from_iso(since)
    until_dt = from_iso(until)
    # Jet filter dates must be 12-hour + AM/PM; %H with %p would emit
    # e.g. "14:30 PM", which Outlook misparses for afternoon times.
    if since_dt:
        clauses.append(f"[ReceivedTime] >= '{since_dt.strftime('%m/%d/%Y %I:%M %p')}'")
    if until_dt:
        clauses.append(f"[ReceivedTime] <= '{until_dt.strftime('%m/%d/%Y %I:%M %p')}'")

    if clauses:
        items = items.Restrict(" AND ".join(clauses))

    from_lower = from_address.lower() if from_address else None
    results: list[dict[str, Any]] = []
    skipped = 0
    for item in items:
        if not _mail_matches_filters(
            item,
            unread_only=unread_only,
            since_dt=since_dt,
            until_dt=until_dt,
            from_address=from_lower,
        ):
            continue
        if skipped < offset:
            skipped += 1
            continue
        results.append(_mail_summary(item))
        if len(results) >= limit:
            break

    return {
        "folder": f.Name,
        "count": len(results),
        "offset": offset,
        "limit": limit,
        "items": results,
        "has_more": len(results) == limit,
        "next_offset": offset + len(results) if len(results) == limit else None,
    }


def search_mails(
    outlook: Any,
    namespace: Any,
    *,
    query: str,
    folder: str | None = "inbox",
    folders: list[str] | None = None,
    limit: int = 25,
    scope: str = "subject_body",
    unread_only: bool = False,
    since: str | None = None,
    until: str | None = None,
    from_address: str | None = None,
    has_attachments: bool | None = None,
    importance: str | None = None,
    categories_contains: list[str] | None = None,
    conversation_id: str | None = None,
) -> dict[str, Any]:
    folder_objs = _resolve_mail_folders(namespace, folder=folder, folders=folders)
    since_dt = from_iso(since)
    until_dt = from_iso(until)
    results: list[tuple[dt.datetime | None, dict[str, Any]]] = []

    simple_single_folder = (
        len(folder_objs) == 1
        and not unread_only
        and since is None
        and until is None
        and from_address is None
        and has_attachments is None
        and importance is None
        and not categories_contains
        and conversation_id is None
    )
    if simple_single_folder:
        f = folder_objs[0]
        items = f.Items
        items.Sort("[ReceivedTime]", True)
        remaining: list[str] = []
        if scope == "dasl":
            filtered = items.Restrict(query)
        else:
            anchor, remaining = split_search_words(query)
            esc = safe_dasl(anchor)
            if scope == "subject":
                filtered = items.Restrict(
                    f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{esc}%'"
                )
            elif scope == "from":
                filtered = items.Restrict(
                    f"@SQL=(\"urn:schemas:httpmail:fromemail\" LIKE '%{esc}%' OR "
                    f"\"urn:schemas:httpmail:fromname\" LIKE '%{esc}%' OR "
                    f"\"{SMTP_PROPTAG}\" LIKE '%{esc}%')"
                )
            else:
                filtered = items.Restrict(
                    f"@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{esc}%' OR "
                    f"\"urn:schemas:httpmail:textdescription\" LIKE '%{esc}%')"
                )
        for item in filtered:
            if not _mail_matches_filters(item, query=query, scope=scope, remaining_words=remaining):
                continue
            results.append((_mail_timestamp(item), _mail_summary(item)))
            if len(results) >= limit:
                break
    else:
        for folder_obj in folder_objs:
            for item in _iter_collection(_safe_get(folder_obj, "Items")):
                if not _mail_matches_filters(
                    item,
                    query=query,
                    scope=scope,
                    unread_only=unread_only,
                    since_dt=since_dt,
                    until_dt=until_dt,
                    from_address=from_address,
                    has_attachments=has_attachments,
                    importance=importance,
                    categories_contains=categories_contains,
                    conversation_id=conversation_id,
                ):
                    continue
                results.append((_mail_timestamp(item), _mail_summary(item)))
        results.sort(
            key=lambda pair: pair[0] or dt.datetime.min,
            reverse=True,
        )
        results = results[:limit]

    items_out = [item for _, item in results]
    return {
        "query": query,
        "scope": scope,
        "folder": folder_objs[0].Name if len(folder_objs) == 1 else None,
        "folders": [_folder_path(folder_obj) for folder_obj in folder_objs],
        "count": len(items_out),
        "items": items_out,
    }


def get_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    include_body: bool = True,
    include_html: bool = False,
    max_body_chars: int = 10000,
) -> dict[str, Any]:
    return _mail_full(
        get_item_by_id(namespace, entry_id),
        include_body=include_body,
        include_html=include_html,
        max_body_chars=max_body_chars,
    )


def send_mail(
    outlook: Any,
    namespace: Any,
    *,
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    html: bool = False,
    attachments: list[str] | None = None,
    importance: str = "normal",
    save_only: bool = False,
    confirm_send: bool = False,
) -> dict[str, Any]:
    mail = outlook.CreateItem(OL_MAIL_ITEM)
    mail.To = "; ".join(to)
    if cc:
        mail.CC = "; ".join(cc)
    if bcc:
        mail.BCC = "; ".join(bcc)
    mail.Subject = subject
    if html:
        mail.BodyFormat = OL_FORMAT_HTML
        mail.HTMLBody = body
    else:
        mail.BodyFormat = OL_FORMAT_PLAIN
        mail.Body = body
    mail.Importance = IMPORTANCE_MAP.get(importance.lower(), OL_IMPORTANCE_NORMAL)

    for raw_path in attachments or []:
        mail.Attachments.Add(validate_attachment_path(raw_path))

    if save_only:
        mail.Save()
        return {
            "status": "saved_to_drafts",
            "entry_id": mail.EntryID,
            "subject": mail.Subject,
        }

    _require_send_confirmation(confirm_send=confirm_send, action="Sending mail")
    mail.Send()
    return {
        "status": "sent",
        "to": to,
        "cc": cc or [],
        "bcc": bcc or [],
        "subject": subject,
    }


def reply_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    body: str,
    reply_all: bool = False,
    html: bool = False,
    attachments: list[str] | None = None,
    save_only: bool = False,
    confirm_send: bool = False,
) -> dict[str, Any]:
    original = get_item_by_id(namespace, entry_id)
    reply = original.ReplyAll() if reply_all else original.Reply()
    if html:
        reply.BodyFormat = OL_FORMAT_HTML
        reply.HTMLBody = body + (reply.HTMLBody or "")
    else:
        reply.Body = body + "\n\n" + (reply.Body or "")
    for raw_path in attachments or []:
        reply.Attachments.Add(validate_attachment_path(raw_path))
    if save_only:
        reply.Save()
        return {
            "status": "saved_to_drafts",
            "entry_id": reply.EntryID,
            "reply_all": reply_all,
            "in_reply_to": entry_id,
            "subject": reply.Subject,
        }
    _require_send_confirmation(confirm_send=confirm_send, action="Replying to mail")
    # Cache properties BEFORE Send(): once the reply is sent, the underlying
    # COM object has effectively moved from Drafts to Sent Items and reading
    # any of its properties raises a "item has been moved or deleted" COM
    # error. Surfacing that as a failure causes upstream AI agents to retry
    # and send duplicates, even though the original Send() succeeded.
    reply_subject = reply.Subject
    reply.Send()
    return {
        "status": "sent",
        "reply_all": reply_all,
        "in_reply_to": entry_id,
        "subject": reply_subject,
    }


def forward_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    to: list[str],
    body: str = "",
    cc: list[str] | None = None,
    html: bool = False,
    save_only: bool = False,
    confirm_send: bool = False,
) -> dict[str, Any]:
    original = get_item_by_id(namespace, entry_id)
    fwd = original.Forward()
    fwd.To = "; ".join(to)
    if cc:
        fwd.CC = "; ".join(cc)
    if body:
        if html:
            fwd.BodyFormat = OL_FORMAT_HTML
            fwd.HTMLBody = body + (fwd.HTMLBody or "")
        else:
            fwd.Body = body + "\n\n" + (fwd.Body or "")
    if save_only:
        fwd.Save()
        return {
            "status": "saved_to_drafts",
            "entry_id": fwd.EntryID,
            "forwarded": entry_id,
            "to": to,
            "subject": fwd.Subject,
        }
    _require_send_confirmation(confirm_send=confirm_send, action="Forwarding mail")
    # Cache properties BEFORE Send(): once the forward is sent, the underlying
    # COM object has effectively moved from Drafts to Sent Items and reading
    # any of its properties raises a "item has been moved or deleted" COM
    # error. Surfacing that as a failure causes upstream AI agents to retry
    # and send duplicates, even though the original Send() succeeded.
    fwd_subject = fwd.Subject
    fwd.Send()
    return {"status": "sent", "forwarded": entry_id, "to": to, "subject": fwd_subject}


def create_draft(
    outlook: Any,
    namespace: Any,
    *,
    to: list[str] | None = None,
    subject: str = "",
    body: str = "",
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    html: bool = False,
    attachments: list[str] | None = None,
    importance: str = "normal",
    categories: str | None = None,
) -> dict[str, Any]:
    payload = send_mail(
        outlook,
        namespace,
        to=to or [],
        subject=subject,
        body=body,
        cc=cc,
        bcc=bcc,
        html=html,
        attachments=attachments,
        importance=importance,
        save_only=True,
        confirm_send=False,
    )
    if categories:
        item = get_item_by_id(namespace, payload["entry_id"])
        item.Categories = categories
        item.Save()
        payload["categories"] = item.Categories or ""
    return payload


def update_draft(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    to: list[str] | None = None,
    subject: str | None = None,
    body: str | None = None,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    html: bool | None = None,
    attachments_to_add: list[str] | None = None,
    clear_attachments: bool = False,
    importance: str | None = None,
    categories: str | None = None,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    if bool(_safe_get(item, "Sent", False)):
        raise OutlookError("This draft has already been sent and can no longer be updated.")
    if to is not None:
        item.To = "; ".join(to)
    if cc is not None:
        item.CC = "; ".join(cc)
    if bcc is not None:
        item.BCC = "; ".join(bcc)
    if subject is not None:
        item.Subject = subject
    if body is not None:
        use_html = bool(html)
        if use_html:
            item.BodyFormat = OL_FORMAT_HTML
            item.HTMLBody = body
        else:
            item.BodyFormat = OL_FORMAT_PLAIN
            item.Body = body
    elif html is not None:
        item.BodyFormat = OL_FORMAT_HTML if html else OL_FORMAT_PLAIN
    if clear_attachments:
        attachments = _safe_get(item, "Attachments")
        while attachments and attachments.Count > 0:
            attachments.Remove(1)
    for raw_path in attachments_to_add or []:
        item.Attachments.Add(validate_attachment_path(raw_path))
    if importance is not None:
        item.Importance = IMPORTANCE_MAP.get(importance.lower(), OL_IMPORTANCE_NORMAL)
    if categories is not None:
        item.Categories = categories
    item.Save()
    return {
        "status": "updated",
        "entry_id": item.EntryID,
        "subject": item.Subject,
        "to": _safe_get(item, "To", ""),
        "cc": _safe_get(item, "CC", ""),
        "bcc": _safe_get(item, "BCC", ""),
        "categories": _safe_get(item, "Categories", "") or "",
        "attachment_count": _safe_get(_safe_get(item, "Attachments"), "Count", 0),
    }


def send_draft(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    confirm_send: bool = False,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    if bool(_safe_get(item, "Sent", False)):
        raise OutlookError("This draft has already been sent.")
    _require_send_confirmation(confirm_send=confirm_send, action="Sending draft")
    subject = _safe_get(item, "Subject", "")
    to = [part.strip() for part in str(_safe_get(item, "To", "") or "").split(";") if part.strip()]
    cc = [part.strip() for part in str(_safe_get(item, "CC", "") or "").split(";") if part.strip()]
    bcc = [part.strip() for part in str(_safe_get(item, "BCC", "") or "").split(";") if part.strip()]
    item.Send()
    return {
        "status": "sent",
        "entry_id": entry_id,
        "subject": subject,
        "to": to,
        "cc": cc,
        "bcc": bcc,
    }


def list_conversation(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str | None = None,
    conversation_id: str | None = None,
    folders: list[str] | None = None,
    limit: int = 100,
) -> dict[str, Any]:
    if not conversation_id:
        if not entry_id:
            raise OutlookError("Provide either entry_id or conversation_id.")
        conversation_id = _safe_get(get_item_by_id(namespace, entry_id), "ConversationID")
    if not conversation_id:
        raise OutlookError("The referenced mail has no conversation_id.")

    folder_objs = _resolve_mail_folders(
        namespace, folders=folders or DEFAULT_CONVERSATION_FOLDERS
    )
    results: list[tuple[dt.datetime | None, dict[str, Any]]] = []
    for folder_obj in folder_objs:
        for item in _iter_collection(_safe_get(folder_obj, "Items")):
            if not _mail_matches_filters(item, conversation_id=conversation_id):
                continue
            results.append((_mail_timestamp(item), _mail_summary(item)))
    results.sort(key=lambda pair: pair[0] or dt.datetime.min)
    items_out = [item for _, item in results[:limit]]
    return {
        "conversation_id": conversation_id,
        "count": len(items_out),
        "folders": [_folder_path(folder_obj) for folder_obj in folder_objs],
        "items": items_out,
    }


def move_mail(outlook: Any, namespace: Any, *, entry_id: str, target_folder: str) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    target = resolve_folder(namespace, target_folder)
    moved = item.Move(target)
    return {"status": "moved", "new_entry_id": moved.EntryID, "folder": target.Name}


def delete_mail(outlook: Any, namespace: Any, *, entry_id: str) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    subject = _safe_get(item, "Subject", "")
    item.Delete()
    return {"status": "deleted", "subject": subject, "entry_id": entry_id}


def mark_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    read: bool | None = None,
    flagged: bool | None = None,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    if read is not None:
        item.UnRead = not read
    if flagged is not None:
        item.FlagStatus = 2 if flagged else 0
    item.Save()
    return {
        "status": "updated",
        "entry_id": entry_id,
        "unread": bool(item.UnRead),
        "flagged": item.FlagStatus == 2,
    }


def save_attachments(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    output_dir: str,
    attachment_index: int | None = None,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    out_dir = validate_output_dir(output_dir)
    saved: list[str] = []
    attachments = list(item.Attachments)
    if attachment_index is not None:
        if attachment_index < 1 or attachment_index > len(attachments):
            raise OutlookError(
                f"attachment_index {attachment_index} out of range "
                f"(message has {len(attachments)} attachments, 1-indexed)."
            )
        attachments = [attachments[attachment_index - 1]]

    for att in attachments:
        # Sender-controlled filename. Reject anything containing path
        # separators, drive-letter prefixes, dot-only names, or reserved
        # Windows device names. These are signals of a path-traversal
        # attempt by the sender, not legitimate attachments.
        raw = att.FileName or ""
        if not raw or raw in (".", ".."):
            raise OutlookError(f"Attachment has invalid filename: {raw!r}")
        if "\\" in raw or "/" in raw:
            raise OutlookError(
                f"Attachment filename contains path separators "
                f"(rejected for safety): {raw!r}"
            )
        if ":" in raw:
            raise OutlookError(
                f"Attachment filename contains colon "
                f"(rejected for safety): {raw!r}"
            )
        # Defense in depth: basename should be a no-op after the checks
        # above, but use it anyway in case ntpath sees something we missed.
        safe_name = ntpath.basename(raw)
        if safe_name != raw:
            raise OutlookError(
                f"Attachment filename did not normalize cleanly: {raw!r}"
            )
        stem = safe_name.lstrip(".").split(".", 1)[0].upper()
        if stem in WINDOWS_RESERVED_DEVICE_NAMES:
            raise OutlookError(
                f"Attachment has reserved Windows device name: {safe_name!r}"
            )

        # Mails often carry several attachments with the same name (e.g.
        # multiple inline "image.png") — uniquify instead of overwriting.
        target = os.path.join(out_dir, safe_name)
        base, ext = os.path.splitext(safe_name)
        counter = 1
        while os.path.exists(target):
            target = os.path.join(out_dir, f"{base} ({counter}){ext}")
            counter += 1
        att.SaveAsFile(target)
        saved.append(target)
    return {
        "status": "saved",
        "count": len(saved),
        "files": saved,
        "output_dir": out_dir,
    }
