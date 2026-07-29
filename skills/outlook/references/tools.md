# Tool reference

Every `outlook_*` tool, with parameters, defaults, return shape, and notes on chaining. Skim the table of contents, then jump to the tools you need.

## Contents

- [Mail](#mail) — list_mails, search_mails, get_mail, get_mails, send_mail, reply_mail, forward_mail, create_draft, update_draft, send_draft, list_conversation, move_mail, delete_mail, mark_mail, save_attachments
- [Folders](#folders) — list_folders, create_folder
- [Calendar](#calendar) — list_events, get_event, create_event, update_event, delete_event, respond_event
- [Contacts](#contacts) — list_contacts, search_contacts, get_contact, resolve_name
- [Tasks](#tasks) — list_tasks, create_task, complete_task
- [Categories](#categories) — list_categories, create_category, update_category, set_category
- [Rules](#rules) — list_rules, toggle_rule, create_rule, update_rule, delete_rule
- [Out-of-Office](#out-of-office) — get_out_of_office
- [Account](#account) — whoami
- [Common return-field glossary](#common-return-field-glossary)

---

## Mail

### `outlook_list_mails`

List mail items from a folder, newest first. Read-only.

| Param            | Type      | Default     | Notes |
| ---------------- | --------- | ----------- | ----- |
| `folder`         | string    | `"inbox"`   | Well-known name or path. See SKILL.md → Folder references. |
| `limit`          | int 1–100 | `25`        | Max items to return. |
| `offset`         | int ≥0    | `0`         | Skip this many before returning. Use with the returned `next_offset` to paginate. |
| `unread_only`    | bool      | `false`     | If true, only unread mails. |
| `since`          | ISO-8601  | `null`      | Lower bound on `ReceivedTime`. |
| `until`          | ISO-8601  | `null`      | Upper bound on `ReceivedTime`. |
| `from_address`   | string    | `null`      | **Substring** match on sender. See gotcha re: `EX:/O=...` addresses. |
| `include_preview`| bool      | `false`     | Include a short body excerpt. Leave disabled for the fastest listing. |
| `response_format`| `markdown`/`json` | `markdown` | Use `json` to extract `entry_id`s. |

**Returns** (`json` shape): `{ folder, count, offset, limit, items: [...], has_more, next_offset }`. Each item has: `entry_id, subject, from, from_address, to, received, unread, has_attachments, importance`; `preview` is present only when requested.

### `outlook_search_mails`

Search one or more folders by subject/body, subject-only, sender, or raw DASL, with optional mail filters. Read-only.

| Param            | Type     | Default          | Notes |
| ---------------- | -------- | ---------------- | ----- |
| `query`          | string   | required         | Search words (ALL must match, any order) — or a DASL @SQL filter when `scope='dasl'`. |
| `folder`         | string   | `"inbox"`        | Where to search. |
| `folders`        | list[str]| `null`           | Search across multiple folders instead of a single folder. |
| `scope`          | enum     | `"subject_body"` | `subject_body` (default), `subject`, `from`, or `dasl`. |
| `limit`          | int 1–100| `25`             | |
| `unread_only`    | bool     | `false`          | Restrict to unread mail. |
| `since`          | ISO-8601 | `null`           | Lower bound on the mail timestamp. |
| `until`          | ISO-8601 | `null`           | Upper bound on the mail timestamp. |
| `from_address`   | string   | `null`           | Sender substring filter, applied in Python across folders. |
| `has_attachments`| bool     | `null`           | Require mail with or without attachments. |
| `importance`     | enum/null| `null`           | `low`, `normal`, or `high`. |
| `categories_contains` | list[str] | `null`     | Require all listed categories. |
| `conversation_id`| string   | `null`           | Restrict to a single Outlook thread. |
| `include_preview`| bool     | `false`          | Include a short body excerpt. Leave disabled for faster searches. |
| `response_format`| str      | `markdown`       | |

**Returns**: `{ query, scope, folder, folders, count, items: [...] }`. Items have the same summary shape as `list_mails`.

Multi-word queries match items containing **all** the words (not the exact phrase), so `"teams not working"` finds "MESP-1 teams is not working". `scope='from'` matches display name, raw address, **and** the real SMTP address (works for Exchange senders too).

`scope='dasl'` is for power use — pass a complete `@SQL=...` filter and the server applies it raw. Only reach for this when subject_body/subject/from can't express what the user wants.

### `outlook_get_mail`

Fetch the body, all headers, and the attachment manifest for one mail. Read-only.

| Param            | Type   | Default | Notes |
| ---------------- | ------ | ------- | ----- |
| `entry_id`       | string | required | From a list/search result. |
| `include_body`   | bool   | `true`   | If false, omits `body`. Useful when you only need metadata. |
| `include_html`   | bool   | `false`  | Adds the raw `html_body`. Usually huge — leave off unless you specifically need the markup. |
| `max_body_chars` | int ≥0 | `10000`  | Body truncation cap; `0` = unlimited. |
| `response_format` | str | `markdown` | |

**Returns**: `{ entry_id, conversation_id, subject, from, from_address, to, cc, bcc, received, sent, unread, importance, categories, attachments: [{index, filename, size_bytes}], body }` plus `body_truncated`/`body_total_chars` when the cap was hit (re-call with a higher `max_body_chars` to read more) and `html_body` when `include_html=true`.

`attachments[].index` is **1-indexed**; pass it to `save_attachments` to save a single file.

### `outlook_get_mails`

Fetch up to 50 messages in one COM queue operation. Prefer this over several
parallel `get_mail` calls because Outlook's object model is STA-only and would
serialize those calls anyway.

| Param            | Type       | Default | Notes |
| ---------------- | ---------- | ------- | ----- |
| `entry_ids`      | list[str]  | required | Between 1 and 50 IDs from list/search results. |
| `include_body`   | bool       | `false` | Include plain-text bodies. |
| `include_html`   | bool       | `false` | Include raw HTML bodies; expensive. |
| `max_body_chars` | int ≥0     | `10000` | Per-message body cap; `0` = unlimited. |
| `response_format`| str        | `"json"` | |

**Returns**: `{ count, items: [...], errors: [{entry_id, error}] }`. One stale
ID does not discard the other successfully retrieved messages.

### `outlook_send_mail`

Compose and send a new mail, or save it to Drafts. Has external side effect.

| Param          | Type      | Default   | Notes |
| -------------- | --------- | --------- | ----- |
| `to`           | list[str] | required  | One or more SMTP addresses. |
| `subject`      | string    | required  | |
| `body`         | string    | required  | Plain text unless `html=true`. |
| `cc`           | list[str] | `null`    | |
| `bcc`          | list[str] | `null`    | |
| `html`         | bool      | `false`   | When true, `body` is HTML. |
| `attachments`  | list[str] | `null`    | Absolute paths under user profile. |
| `importance`   | enum      | `"normal"`| `low` / `normal` / `high`. |
| `save_only`    | bool      | `false`   | **Save to Drafts instead of sending.** |
| `confirm_send` | bool      | `false`   | Required to actually send when `save_only=false`. |

**Returns** (sent): `{ status: "sent", to, cc, bcc, subject }`. (Drafts): `{ status: "saved_to_drafts", entry_id, subject }`.

Always confirm the recipient list and subject with the user before calling this tool unless they have explicitly authorized you to send.

### `outlook_reply_mail`

Reply (or reply-all) to an existing mail. The original message is appended below your body, the same way Outlook's Reply button does it. Has external side effect.

| Param         | Type      | Default | Notes |
| ------------- | --------- | ------- | ----- |
| `entry_id`    | string    | required | The mail being replied to. |
| `body`        | string    | required | Your reply text. The quoted original is appended automatically. |
| `reply_all`   | bool      | `false`  | If true, includes the original CC list. |
| `html`        | bool      | `false`  | |
| `attachments` | list[str] | `null`   | |
| `save_only`   | bool      | `false`  | Save the reply to Drafts instead of sending. |
| `confirm_send`| bool      | `false`  | Required to actually send when `save_only=false`. |

**Returns**: sent: `{ status: "sent", reply_all, in_reply_to, subject }`; draft: `{ status: "saved_to_drafts", entry_id, reply_all, in_reply_to, subject }`.

### `outlook_forward_mail`

Forward an existing mail to new recipients with an optional note above. Has external side effect.

| Param      | Type      | Default | Notes |
| ---------- | --------- | ------- | ----- |
| `entry_id` | string    | required | |
| `to`       | list[str] | required | |
| `body`     | string    | `""`     | Optional note prepended to the forwarded content. |
| `cc`       | list[str] | `null`   | |
| `html`     | bool      | `false`  | |
| `save_only`   | bool      | `false`  | Save the forward to Drafts instead of sending. |
| `confirm_send`| bool      | `false`  | Required to actually send when `save_only=false`. |

**Returns**: sent: `{ status: "sent", forwarded, to, subject }`; draft: `{ status: "saved_to_drafts", entry_id, forwarded, to, subject }`.

### `outlook_create_draft`

Create a new draft directly in Drafts.

| Param         | Type      | Default | Notes |
| ------------- | --------- | ------- | ----- |
| `to`          | list[str] | `null`  | Optional recipients. |
| `subject`     | string    | `""`    | |
| `body`        | string    | `""`    | |
| `cc`          | list[str] | `null`  | |
| `bcc`         | list[str] | `null`  | |
| `html`        | bool      | `false` | |
| `attachments` | list[str] | `null`  | |
| `importance`  | enum      | `"normal"` | `low` / `normal` / `high`. |
| `categories`  | string    | `null`  | Optional comma-separated categories. |

**Returns**: `{ status: "saved_to_drafts", entry_id, subject }` plus `categories` when set.

### `outlook_update_draft`

Update an existing draft in place.

| Param               | Type      | Default | Notes |
| ------------------- | --------- | ------- | ----- |
| `entry_id`          | string    | required | Draft EntryID. |
| `to`                | list[str] | `null`  | Replace To list. |
| `subject`           | string    | `null`  | Replace subject. |
| `body`              | string    | `null`  | Replace body. |
| `cc`                | list[str] | `null`  | Replace CC list. |
| `bcc`               | list[str] | `null`  | Replace BCC list. |
| `html`              | bool      | `null`  | If `body` is present, controls body format. |
| `attachments_to_add`| list[str] | `null`  | Append files. |
| `clear_attachments` | bool      | `false` | Remove all existing attachments first. |
| `importance`        | enum/null | `null`  | Replace importance. |
| `categories`        | string    | `null`  | Replace comma-separated categories. |

**Returns**: `{ status: "updated", entry_id, subject, to, cc, bcc, categories, attachment_count }`.

### `outlook_send_draft`

Send a previously saved draft.

| Param          | Type   | Default | Notes |
| -------------- | ------ | ------- | ----- |
| `entry_id`     | string | required | Draft EntryID. |
| `confirm_send` | bool   | `false` | Required to actually send the draft. |

### `outlook_list_conversation`

List one Outlook thread across common folders or a custom folder set.

| Param             | Type      | Default | Notes |
| ----------------- | --------- | ------- | ----- |
| `entry_id`        | string    | `null`  | Seed mail EntryID. Use this or `conversation_id`. |
| `conversation_id` | string    | `null`  | Explicit Outlook conversation id. |
| `folders`         | list[str] | `null`  | Defaults to inbox, sent, drafts, deleted. |
| `limit`           | int 1–200 | `100`   | Max items returned. |
| `response_format` | str       | `markdown` | |

**Returns**: `{ conversation_id, count, folders, items: [...] }`, oldest-first within the conversation.

### `outlook_move_mail`

Move a mail to another folder.

| Param           | Type   | Default | Notes |
| --------------- | ------ | ------- | ----- |
| `entry_id`      | string | required | |
| `target_folder` | string | required | Well-known name or path. |

**Returns**: `{ status: "moved", new_entry_id, folder }`.

The `entry_id` changes when an item moves stores. **Use the returned `new_entry_id`** if you need to act on the moved item again.

### `outlook_delete_mail`

Soft-delete (moves to Deleted Items). Reversible by the user from Outlook.

| Param      | Type   | Default | Notes |
| ---------- | ------ | ------- | ----- |
| `entry_id` | string | required | |

**Returns**: `{ status: "deleted", subject, entry_id }`.

### `outlook_mark_mail`

Toggle read state and/or follow-up flag.

| Param      | Type | Default | Notes |
| ---------- | ---- | ------- | ----- |
| `entry_id` | string | required | |
| `read`     | bool/null | `null` | `true` = mark read, `false` = mark unread, `null` = no change. |
| `flagged`  | bool/null | `null` | `true` = flag for follow-up, `false` = clear flag, `null` = no change. |

**Returns**: `{ status: "updated", entry_id, unread, flagged }`.

### `outlook_save_attachments`

Save one or all attachments from a mail to a local directory.

| Param              | Type    | Default | Notes |
| ------------------ | ------- | ------- | ----- |
| `entry_id`         | string  | required | |
| `output_dir`       | string  | required | Absolute path under user profile. Created if missing. |
| `attachment_index` | int ≥1  | `null`  | 1-indexed. Omit to save all. |

**Returns**: `{ status: "saved", count, files: [absolute paths], output_dir }`.

---

## Folders

### `outlook_list_folders`

Walk the folder tree under a root.

| Param            | Type    | Default | Notes |
| ---------------- | ------- | ------- | ----- |
| `root`           | string  | `null`  | Folder to start from. Default = the default mailbox root. |
| `max_depth`      | int 1–10| `4`     | How deep to walk. |
| `response_format`| str     | `markdown` | |

**Returns**: `{ count, items: [{name, path, item_count, unread_count, default_item_type}, ...] }`. The `path` strings are exactly what you pass back as a `folder` parameter elsewhere.

### `outlook_create_folder`

Create a sub-folder under a parent.

| Param    | Type   | Default   | Notes |
| -------- | ------ | --------- | ----- |
| `name`   | string | required  | New folder name. |
| `parent` | string | `"inbox"` | Where to put it. |

**Returns**: `{ name, path, entry_id }`.

---

## Calendar

### `outlook_list_events`

List calendar events in a date range, including expanded recurring instances.

| Param                 | Type    | Default        | Notes |
| --------------------- | ------- | -------------- | ----- |
| `start`               | ISO-8601| now            | |
| `end`                 | ISO-8601| `start + 14d`  | |
| `limit`               | int 1–200| `50`          | |
| `include_recurrences` | bool    | `true`         | If false, only the master entries — usually you want true. |
| `response_format`     | str     | `markdown`     | |

**Returns**: `{ start, end, count, items: [...] }`. Items: `entry_id, subject, start, end, location, organizer, is_recurring, all_day, preview` (200-char body excerpt).

### `outlook_get_event`

Full event detail, including attendees and their RSVP status.

| Param      | Type   | Default | Notes |
| ---------- | ------ | ------- | ----- |
| `entry_id` | string | required | |
| `response_format` | str | `markdown` | |

**Returns**: summary fields + `body, attendees: [{name, address, type, response}], reminder_minutes, categories`.

### `outlook_create_event`

Create a calendar event or meeting invite. **Adding any attendee turns this into a meeting that is sent immediately on success — there is no draft state for meeting invites.**

| Param               | Type           | Default | Notes |
| ------------------- | -------------- | ------- | ----- |
| `subject`           | string         | required | |
| `start`             | ISO-8601       | required | |
| `end`               | ISO-8601       | required | |
| `location`          | string         | `null`  | |
| `body`              | string         | `null`  | |
| `attendees`         | list[str]      | `null`  | Email addresses. Adding any value here makes this a meeting and sends invites. |
| `is_online_meeting` | bool           | `false` | Reserved — current behavior is to mark the meeting; the actual Teams/Zoom link is added by Outlook clients. |
| `reminder_minutes`  | int 0–10080    | `15`    | Minutes before start. |
| `recurrence`        | Recurrence obj | `null`  | See SKILL.md → Recurrence. |

**Returns**: `{ status: "created", entry_id, subject, start, end }`.

Confirm attendee list, times, and recurrence with the user before calling.

### `outlook_update_event`

Update fields on an event. Only non-null fields are written. Does **not** modify recurrence — for cadence changes, delete and recreate.

| Param      | Type     | Default | Notes |
| ---------- | -------- | ------- | ----- |
| `entry_id` | string   | required | |
| `subject`  | string   | `null`  | |
| `start`    | ISO-8601 | `null`  | |
| `end`      | ISO-8601 | `null`  | |
| `location` | string   | `null`  | |
| `body`     | string   | `null`  | |

**Returns**: `{ status: "updated", entry_id }`. If the event has attendees, Outlook may send an updated-meeting notification when this saves.

### `outlook_delete_event`

Delete a calendar event. **If the event has attendees, this sends a cancellation notice.**

**Returns**: `{ status: "deleted", subject, entry_id }`.

### `outlook_respond_event`

Respond to a meeting invite.

| Param           | Type   | Default | Notes |
| --------------- | ------ | ------- | ----- |
| `entry_id`      | string | required | |
| `response`      | enum   | required | `accept` / `tentative` / `decline`. |
| `send_response` | bool   | `true`   | Set false to record locally without emailing the organizer. |

**Returns**: `{ status: "responded", response }`.

---

## Contacts

### `outlook_list_contacts`

List saved contacts from **every contact folder in every store**, sorted by full name within each folder.

| Param            | Type     | Default | Notes |
| ---------------- | -------- | ------- | ----- |
| `limit`          | int 1–200| `50`    | |
| `offset`         | int ≥0   | `0`     | |
| `response_format`| str      | `markdown` | |

**Returns**: `{ count, offset, items: [...], has_more }`. Items: `entry_id, full_name, email, company, job_title, mobile, folder`.

On corporate accounts the personal contact folders are often nearly empty — colleagues live in the **directory** (Global Address List), which this tool does not list. Use `search_contacts` to find people.

### `outlook_search_contacts`

Word search across saved contacts (name, email, company, job title) **and the org directory (GAL)**. All query words must match.

| Param               | Type    | Default | Notes |
| ------------------- | ------- | ------- | ----- |
| `query`             | string  | required | e.g. `"anas shaikh"` matches "Anas Ahmed Shaikh". |
| `limit`             | int 1–100| `25`   | |
| `include_directory` | bool    | `true`  | Also scan the Exchange Global Address List. A few seconds on large directories. |
| `response_format`   | str     | `markdown` | |

**Returns**: `{ query, count, items, searched_directory }`. Each item has `source: "contacts" | "directory"`. Directory items carry `full_name, email (SMTP), company, job_title, mobile` but **no `entry_id`** — they aren't Outlook items, so don't pass them to `get_contact`.

### `outlook_get_contact`

Full contact record (saved contacts only — needs an `entry_id`). Returns the summary fields plus `business_phone, home_phone, address, notes`.

### `outlook_resolve_name`

Resolve a display name, alias, or address to its SMTP address — same mechanism as typing a name in To: and pressing Ctrl+K. Use before sending when you only know a person's name.

| Param  | Type   | Default | Notes |
| ------ | ------ | ------- | ----- |
| `name` | string | required | Full names resolve best; short fragments are often ambiguous. |

**Returns**: `{ resolved: true, query, display_name, smtp_address }` or `{ resolved: false, query, note }`. Ambiguous names do **not** resolve — fall back to `search_contacts` to browse candidates.

---

## Tasks

### `outlook_list_tasks`

List tasks from the default Tasks folder, sorted by due date.

| Param                | Type     | Default | Notes |
| -------------------- | -------- | ------- | ----- |
| `limit`              | int 1–200| `50`    | |
| `include_completed`  | bool     | `false` | Default hides done tasks. |
| `response_format`    | str      | `markdown` | |

**Items**: `entry_id, subject, due_date, start_date, complete, percent_complete, importance, status`.

### `outlook_create_task`

| Param        | Type     | Default   | Notes |
| ------------ | -------- | --------- | ----- |
| `subject`    | string   | required  | |
| `due_date`   | ISO-8601 | `null`    | |
| `body`       | string   | `null`    | |
| `importance` | enum     | `"normal"`| low/normal/high. |
| `reminder`   | ISO-8601 | `null`    | Sets a reminder time. |

**Returns**: `{ status: "created", entry_id, subject }`.

### `outlook_complete_task`

| Param      | Type   | Default | Notes |
| ---------- | ------ | ------- | ----- |
| `entry_id` | string | required | |

Marks the task 100% complete. Returns `{ status: "completed", entry_id }`.

---

## Categories

### `outlook_list_categories`

Returns the color categories defined in the user's Outlook profile: `{ count, items: [{name, color}, ...] }`. Categories are profile-wide, not per-folder.

### `outlook_create_category`

Creates a new category in the master category list for this Outlook profile.

| Param   | Type        | Default | Notes |
| ------- | ----------- | ------- | ----- |
| `name`  | string      | required | Category display name. |
| `color` | int / null  | `null`  | Optional Outlook color enum value from 0 to 25. Omit to let Outlook choose automatically. |

**Returns**: `{ status: "created", category: {name, color} }`.

### `outlook_update_category`

Renames and/or recolors an existing category in the profile-wide master category list.

| Param      | Type       | Default  | Notes |
| ---------- | ---------- | -------- | ----- |
| `name`     | string     | required | Exact current name from `outlook_list_categories`. |
| `new_name` | string/null | `null`  | Optional replacement name; must not duplicate another category. |
| `color`    | int/null   | `null`   | Optional replacement Outlook color enum value from 0 to 25. |

At least one of `new_name` or `color` is required.

**Returns**: `{ status: "updated", previous_name, category: {name, color} }`.

### `outlook_set_category`

Replace the categories on any item (mail, event, task).

| Param        | Type   | Default | Notes |
| ------------ | ------ | ------- | ----- |
| `entry_id`   | string | required | |
| `categories` | string | required | **Comma-separated names.** Empty string clears all. e.g. `"Important"` or `"Work, Follow-up"`. |

This **replaces** existing categories rather than adding to them. To add `Foo` to an item that already has `Bar`, send `"Bar, Foo"`. Get the current value first via `get_mail` / `get_event` if needed.

---

## Rules

### `outlook_list_rules`

Returns the user's mail rules with their on/off state plus the supported editable subset:
`{ count, items: [{index, name, enabled, execution_order, rule_type, supported_conditions, supported_exceptions, supported_actions}] }`.

### `outlook_toggle_rule`

| Param       | Type   | Default | Notes |
| ----------- | ------ | ------- | ----- |
| `rule_name` | string | required | **Exact** name from `list_rules`. |
| `enabled`   | bool   | required | `true` to enable, `false` to disable. |

This change is live the moment it's saved — no staging buffer. Confirm the rule name with the user before calling.

### `outlook_create_rule`

Creates a **receive rule** using the supported COM-safe subset: sender-address, recipient/alias, current-mailbox-in-To, subject, and body conditions plus move/copy/assign-category actions.

| Param                     | Type         | Default | Notes |
| ------------------------- | ------------ | ------- | ----- |
| `name`                    | string       | required | Display name for the rule. |
| `sender_address_contains` | list[string] | `null`  | Optional OR-match sender substrings. |
| `sent_to_recipients`      | list[string] | `null`  | Optional recipients in To or Cc; names, aliases, or SMTP addresses must resolve in Outlook. |
| `to_me`                   | bool         | `null`  | Require the current mailbox to appear in To. |
| `subject_contains`        | list[string] | `null`  | Optional OR-match subject substrings. |
| `body_contains`           | list[string] | `null`  | Optional OR-match body substrings. |
| `move_to_folder`          | string       | `null`  | Optional move target folder path/name. |
| `copy_to_folder`          | string       | `null`  | Optional copy target folder path/name. |
| `assign_categories`       | list[string] | `null`  | Optional category names to assign. |
| `except_sender_address_contains` | list[string] | `null` | Sender substrings that block the rule. |
| `except_sent_to_recipients` | list[string] | `null` | Recipients in To or Cc that block the rule. |
| `except_to_me`            | bool         | `null`  | Block the rule when the current mailbox appears in To. |
| `except_subject_contains` | list[string] | `null`  | Subject substrings that block the rule. |
| `except_body_contains`    | list[string] | `null`  | Body substrings that block the rule. |
| `stop_processing_more_rules` | bool      | `false` | Stop later rules after this one matches. |
| `execution_order`         | int/null     | `null`  | Optional rule position in the collection. |
| `enabled`                 | bool         | `true`  | Whether the new rule starts enabled. |

At least one supported condition and one supported action are required.

Example for "sent to the administrative alias, except when I am directly in To":

```json
{
  "name": "Administratif ABM",
  "sent_to_recipients": ["administratif@example.com"],
  "except_to_me": true,
  "move_to_folder": "Inbox/Administratif"
}
```

### `outlook_update_rule`

Updates a rule's supported editable fields in place.

| Param                     | Type         | Default | Notes |
| ------------------------- | ------------ | ------- | ----- |
| `rule_name`               | string       | required | **Exact** current rule name from `list_rules`. |
| `new_name`                | string       | `null`  | Rename the rule. |
| `enabled`                 | bool         | `null`  | Toggle on/off; omit to leave unchanged. |
| `sender_address_contains` | list[string] | `null`  | Replace sender substrings; pass `[]` to clear. |
| `sent_to_recipients`      | list[string] | `null`  | Replace recipients matched in To or Cc; pass `[]` to clear. |
| `to_me`                   | bool         | `null`  | Enable/disable the current-mailbox-in-To condition. |
| `subject_contains`        | list[string] | `null`  | Replace subject substrings; pass `[]` to clear. |
| `body_contains`           | list[string] | `null`  | Replace body substrings; pass `[]` to clear. |
| `move_to_folder`          | string       | `null`  | Replace move target folder. |
| `copy_to_folder`          | string       | `null`  | Replace copy target folder. |
| `assign_categories`       | list[string] | `null`  | Replace assigned categories; pass `[]` to clear. |
| `except_sender_address_contains` | list[string] | `null` | Replace sender exceptions; pass `[]` to clear. |
| `except_sent_to_recipients` | list[string] | `null` | Replace recipient exceptions; pass `[]` to clear. |
| `except_to_me`            | bool         | `null`  | Enable/disable the current-mailbox-in-To exception. |
| `except_subject_contains` | list[string] | `null`  | Replace subject exceptions; pass `[]` to clear. |
| `except_body_contains`    | list[string] | `null`  | Replace body exceptions; pass `[]` to clear. |
| `clear_move_to_folder`    | bool         | `false` | Disable the move action entirely. |
| `clear_copy_to_folder`    | bool         | `false` | Disable the copy action entirely. |
| `clear_assign_categories` | bool         | `false` | Disable category assignment entirely. |
| `stop_processing_more_rules` | bool/null | `null`  | Enable/disable stop-processing. |
| `execution_order`         | int/null     | `null`  | Reposition the rule in the collection. |

The rule must still have at least one supported condition and one supported action after the update.

### `outlook_delete_rule`

Delete a rule by exact name.

| Param       | Type   | Default | Notes |
| ----------- | ------ | ------- | ----- |
| `rule_name` | string | required | **Exact** current rule name from `list_rules`. |

---

## Out-of-Office

### `outlook_get_out_of_office`

Reports whether OOO auto-reply is currently on. Returns `{ out_of_office: bool, status: "on"|"off" }`, or `{ out_of_office: null, status: "unknown", note: ... }` if the property isn't readable on this profile.

There is **no tool to enable, disable, or schedule OOO**. Tell users to manage it via Outlook → File → Automatic Replies.

---

## Account

### `outlook_whoami`

Returns the bound user, the account list, and the user's timezone: `{ current_user, accounts: [{display_name, smtp_address, user_name, account_type}, ...], local_time, timezone, utc_offset }`. Useful as a sanity check when the user has multiple mailboxes, and as the authority on what timezone all returned datetimes are in.

---

## Common return-field glossary

- `entry_id` — opaque, stable handle for an item. Pass back verbatim. Becomes invalid on delete; changes on cross-store move.
- `conversation_id` — groups mails in a thread. Same value across replies/forwards in one conversation.
- `from` — display name of the sender.
- `from_address` — sender address. **For Exchange senders, this is an `EX:/O=...` distinguished name, not SMTP.** Match by substring.
- `received` / `sent` / `start` / `end` / `due_date` — ISO-8601 strings **in the user's local timezone with explicit offset** (e.g. `2026-06-10T16:33:22+05:00`). Present as-is; never convert to another timezone.
- `unread` — bool. Note `mark_mail` returns `unread` (not `read`).
- `importance` — integer (0=low, 1=normal, 2=high).
- `categories` — comma-separated string of category names; empty string = none.
- `preview` — optional 200-char body excerpt when `include_preview=true`. Not a substitute for `get_mail` / `get_mails` / `get_event` when you need the full body.
- `has_more` / `next_offset` — pagination signals on list endpoints.
