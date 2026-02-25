# Messages AppleScript Dictionary

> Auto-generated from `Messages.sdef` inside the app bundle.  
> Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Messages"`

## Table of Contents

- [Messages Suite](#messages-suite)

---

## Messages Suite

commands and classes for Messages scripting.

### Commands

### `send`

Sends a message to a participant or to a chat.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `file | text` | no |  |
| `to` | `participant | chat` | no |  |

### `login`

Login to all accounts.

### `logout`

Logout of all accounts.

### Classes

### `participant`

A participant for an account.

*Plural:* `participants`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | The participant's unique identifier. For example: 01234567-89AB-CDEF-0123-456789ABCDEF:+11234567890 |
| `account` | `account` | r/o | The account for this participant. |
| `name` | `text` | r/o | The participant's name as it appears in the participant list. |
| `handle` | `text` | r/o | The participant's handle. |
| `first name` | `text` | r/o | The first name from this participan's Contacts card, if available |
| `last name` | `text` | r/o | The last name from this participant's Contacts card, if available |
| `full name` | `text` | r/o | The full name from this participant's Contacts card, if available |

### `account`

An account that can be logged in to from this system

*Plural:* `accounts`

**Contains:** `chat`, `participant`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | A unique identifier for this account. |
| `description` | `text` | r/o | The name of this account as defined in Account preferences description field |
| `enabled` | `boolean` | r/w | Is the account enabled? |
| `connection status` | `connection status` | r/o | The connection status for this account. |
| `service type` | `service type` | r/o | The type of service for this account |

### `chat`

An SMS or iMessage chat.

**Contains:** `participant`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | A guid identifier for this chat. |
| `name` | `text` | r/o | The chat's name as it appears in the chat list. |
| `account` | `account` | r/o | The account which is participating in this chat. |

### `file transfer`

A file being sent or received

*Plural:* `file transfers`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | The guid for this file transfer |
| `name` | `text` | r/o | The name of this file |
| `file path` | `file` | r/o | The local path to this file transfer |
| `direction` | `direction` | r/o | The direction in which this file is being sent |
| `account` | `account` | r/o | The account on which this file transfer is taking place |
| `participant` | `participant` | r/o | The other participatant in this file transfer |
| `file size` | `integer` | r/o | The total size in bytes of the completed file transfer |
| `file progress` | `integer` | r/o | The number of bytes that have been transferred |
| `transfer status` | `transfer status` | r/o | The current status of this file transfer |
| `started` | `date` | r/o | The date that this file transfer started |

### Enumerations

### `service type`

| Value | Description |
|-------|-------------|
| `SMS` |  |
| `iMessage` |  |
| `RCS` |  |

### `direction`

| Value | Description |
|-------|-------------|
| `incoming` |  |
| `outgoing` |  |

### `transfer status`

| Value | Description |
|-------|-------------|
| `preparing` |  |
| `waiting` |  |
| `transferring` |  |
| `finalizing` |  |
| `finished` |  |
| `failed` |  |

### `connection status`

| Value | Description |
|-------|-------------|
| `disconnecting` |  |
| `connected` |  |
| `connecting` |  |
| `disconnected` |  |
