# Mail AppleScript Dictionary

> Auto-generated from `Mail.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Mail"`

## Table of Contents

- [Standard Suite](#standard-suite)
- [Text Suite](#text-suite)
- [Mail](#mail)
- [Mail Framework](#mail-framework)

---

## Standard Suite

Common classes and commands for all applications.

### Commands

### `delete`

Delete an object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The object(s) to delete. |

### `duplicate`

Copy an object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The object(s) to copy. |
| `to` | `location specifier` | yes | The location for the new copy or copies. |
| `with properties` | `record` | yes | Properties to set in the new copy or copies right away. |

### `move`

Move an object to a new location.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | The object(s) to move. |
| `to` | `location specifier` | no | The new location for the object(s). |

---

## Text Suite

A set of basic classes for text processing.

### Classes

### `rich text`

Rich (styled) text

*Plural:* `rich text`

**Contains:** `paragraph`, `word`, `character`, `attribute run`, `attachment`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color` | `RGB color` | r/w | The color of the first character. |
| `font` | `text` | r/w | The name of the font of the first character. |
| `size` | `number` | r/w | The size in points of the first character. |

### `attachment`

Represents an inline text attachment. This class is used mainly for make commands.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `file name` | `file` | r/w | The file for the attachment |

### `paragraph`

This subdivides the text into paragraphs.

**Contains:** `word`, `character`, `attribute run`, `attachment`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color` | `RGB color` | r/w | The color of the first character. |
| `font` | `text` | r/w | The name of the font of the first character. |
| `size` | `number` | r/w | The size in points of the first character. |

### `word`

This subdivides the text into words.

**Contains:** `character`, `attribute run`, `attachment`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color` | `RGB color` | r/w | The color of the first character. |
| `font` | `text` | r/w | The name of the font of the first character. |
| `size` | `number` | r/w | The size in points of the first character. |

### `character`

This subdivides the text into characters.

**Contains:** `attribute run`, `attachment`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color` | `RGB color` | r/w | The color of the character. |
| `font` | `text` | r/w | The name of the font of the character. |
| `size` | `number` | r/w | The size in points of the character. |

### `attribute run`

This subdivides the text into chunks that all have the same attributes.

**Contains:** `paragraph`, `word`, `character`, `attachment`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color` | `RGB color` | r/w | The color of the first character. |
| `font` | `text` | r/w | The name of the font of the first character. |
| `size` | `number` | r/w | The size in points of the first character. |

---

## Mail

Classes and commands for the Mail application

### Commands

### `bounce`

Does nothing at all (deprecated)

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `message` | no | the message to bounce |

### `check for new mail`

Triggers a check for email.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `for` | `account` | yes | Specify the account that you wish to check for mail |

### `extract name from`

Command to get the full name out of a fully specified email address. E.g. Calling this with "John Doe <jdoe@example.com>" as the direct object would return "John Doe"

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | fully formatted email address |

**Returns:** `text` — the full name

### `extract address from`

Command to get just the email address of a fully specified email address. E.g. Calling this with "John Doe <jdoe@example.com>" as the direct object would return "jdoe@example.com"

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | fully formatted email address |

**Returns:** `text` — the email address

### `forward`

Creates a forwarded message.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `message` | no | the message to forward |
| `opening window` | `boolean` | yes | Whether the window for the forwarded message is shown. Default is to not show the window. |

**Returns:** `outgoing message` — the message to be forwarded

### `GetURL`

Opens a mailto URL.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | the mailto URL |

### `import Mail mailbox`

Imports a mailbox created by Mail.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `at` | `file` | no | the mailbox or folder of mailboxes to import |

### `mailto`

Opens a mailto URL.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | the mailto URL |

### `perform mail action with messages`

Script handler invoked by rules and menus that execute AppleScripts. The direct parameter of this handler is a list of messages being acted upon.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `message` | no | the message being acted upon |
| `in mailboxes` | `mailbox` | yes | If the script is being executed by the user selecting an item in the scripts menu, this argument will specify the mailboxes that are currently selected. Otherwise it will not be specified. |
| `for rule` | `rule` | yes | If the script is being executed by a rule action, this argument will be the rule being invoked. Otherwise it will not be specified. |

### `redirect`

Creates a redirected message.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `message` | no | the message to redirect |
| `opening window` | `boolean` | yes | Whether the window for the redirected message is shown. Default is to not show the window. |

**Returns:** `outgoing message` — the redirected message

### `reply`

Creates a reply message.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `message` | no | the message to reply to |
| `opening window` | `boolean` | yes | Whether the window for the reply message is shown. Default is to not show the window. |
| `reply to all` | `boolean` | yes | Whether to reply to all recipients. Default is to reply to the sender only. |

**Returns:** `outgoing message` — the reply message

### `send`

Sends a message.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `outgoing message` | no | the message to send |

**Returns:** `boolean` — true if sending was successful, false if not

### `synchronize`

Command to trigger synchronizing of an IMAP account with the server.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `with` | `account` | no | The account to synchronize |

### Classes

### `outgoing message`

A new email message

**Contains:** `bcc recipient`, `cc recipient`, `recipient`, `to recipient`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `sender` | `text` | r/w | The sender of the message |
| `subject` | `text` | r/w | The subject of the message |
| `visible` | `boolean` | r/w | Controls whether the message window is shown on the screen. The default is false |
| `message signature` | `signature | missing value` | r/w | The signature of the message |
| `id` | `integer` | r/o | The unique identifier of the message |
| `html content` | `text` | w/o | Does nothing at all (deprecated) |
| `vcard path` | `file` | w/o | Does nothing at all (deprecated) |

**Responds to:** `save`, `close`, `send`

### `ldap server`

DEPRECATED - DO NOT USE

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `enabled` | `boolean` | r/w | Does nothing at all (deprecated) |
| `name` | `text` | r/w | Does nothing at all (deprecated) |
| `port` | `integer` | r/w | Does nothing at all (deprecated) |
| `scope` | `LdapScope` | r/w | Does nothing at all (deprecated) |
| `search base` | `text` | r/w | Does nothing at all (deprecated) |
| `host name` | `text` | r/w | Does nothing at all (deprecated) |

### `OLD message editor`

DEPRECATED - DO NOT USE

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `OLD compose message` | `outgoing message` | r/w | DEPRECATED - DO NOT USE |

### `message viewer`

Represents the object responsible for managing a viewer window

**Contains:** `message`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `drafts mailbox` | `mailbox` | r/o | The top level Drafts mailbox |
| `inbox` | `mailbox` | r/o | The top level In mailbox |
| `junk mailbox` | `mailbox` | r/o | The top level Junk mailbox |
| `outbox` | `mailbox` | r/o | The top level Out mailbox |
| `sent mailbox` | `mailbox` | r/o | The top level Sent mailbox |
| `trash mailbox` | `mailbox` | r/o | The top level Trash mailbox |
| `sort column` | `ViewerColumns` | r/w | The column that is currently sorted in the viewer |
| `sorted ascending` | `boolean` | r/w | Whether the viewer is sorted ascending or not |
| `mailbox list visible` | `boolean` | r/w | Controls whether the list of mailboxes is visible or not |
| `preview pane is visible` | `boolean` | r/w | Controls whether the preview pane of the message viewer window is visible or not |
| `visible columns` | `ViewerColumns` | r/w | List of columns that are visible. The subject column and the message status column will always be visible |
| `id` | `integer` | r/o | The unique identifier of the message viewer |
| `visible messages` | `message` | r/w | List of messages currently being displayed in the viewer |
| `selected messages` | `message` | r/w | List of messages currently selected |
| `selected mailboxes` | `mailbox` | r/w | List of mailboxes currently selected in the list of mailboxes |

### `signature`

Email signatures

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `content` | `text` | r/w | Contents of email signature. If there is a version with fonts and/or styles, that will be returned over the plain text version |
| `name` | `text` | r/w | Name of the signature |

### Enumerations

### `saveable file format`

| Value | Description |
|-------|-------------|
| `native format` | Native format |

### `DefaultMessageFormat`

| Value | Description |
|-------|-------------|
| `plain format` | Plain Text |
| `rich format` | Rich Text |

### `HeaderDetail`

| Value | Description |
|-------|-------------|
| `all` | All |
| `custom` | Custom |
| `default` | Default |
| `no headers` | No headers |

### `LdapScope`

| Value | Description |
|-------|-------------|
| `base` | LDAP scope of 'Base' |
| `one level` | LDAP scope of 'One Level' |
| `subtree` | LDAP scope of 'Subtree' |

### `QuotingColor`

| Value | Description |
|-------|-------------|
| `blue` | Blue |
| `green` | Green |
| `orange` | Orange |
| `other` | Other |
| `purple` | Purple |
| `red` | Red |
| `yellow` | Yellow |

### `ViewerColumns`

| Value | Description |
|-------|-------------|
| `attachments column` | Column containing the number of attachments a message contains |
| `message color` | Used to indicate sorting should be done by color |
| `date received column` | Column containing the date a message was received |
| `date sent column` | Column containing the date a message was sent |
| `flags column` | Column containing the flags of a message |
| `from column` | Column containing the sender's name |
| `mailbox column` | Column containing the name of the mailbox or account a message is in |
| `message status column` | Column indicating a messages status (read, unread, replied to, forwarded, etc) |
| `number column` | Column containing the number of a message in a mailbox |
| `size column` | Column containing the size of a message |
| `subject column` | Column containing the subject of a message |
| `to column` | Column containing the recipients of a message |
| `date last saved column` | Column containing the date a draft message was saved |

---

## Mail Framework

Classes and commands for the Mail framework

### Classes

### `message`

An email message

**Contains:** `bcc recipient`, `cc recipient`, `recipient`, `to recipient`, `header`, `mail attachment`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `integer` | r/o | The unique identifier of the message. |
| `all headers` | `text` | r/o | All the headers of the message |
| `background color` | `HighlightColors` | r/w | The background color of the message |
| `mailbox` | `mailbox` | r/w | The mailbox in which this message is filed |
| `content` | `rich text` | r/o | Contents of an email message |
| `date received` | `date` | r/o | The date a message was received |
| `date sent` | `date` | r/o | The date a message was sent |
| `deleted status` | `boolean` | r/w | Indicates whether the message is deleted or not |
| `flagged status` | `boolean` | r/w | Indicates whether the message is flagged or not |
| `flag index` | `integer` | r/w | The flag on the message, or -1 if the message is not flagged |
| `junk mail status` | `boolean` | r/w | Indicates whether the message has been marked junk or evaluated to be junk by the junk mail filter. |
| `read status` | `boolean` | r/w | Indicates whether the message is read or not |
| `message id` | `text` | r/o | The unique message ID string |
| `source` | `text` | r/o | Raw source of the message |
| `reply to` | `text` | r/o | The address that replies should be sent to |
| `message size` | `integer` | r/o | The size (in bytes) of a message |
| `sender` | `text` | r/o | The sender of the message |
| `subject` | `text` | r/o | The subject of the message |
| `was forwarded` | `boolean` | r/o | Indicates whether the message was forwarded or not |
| `was redirected` | `boolean` | r/o | Indicates whether the message was redirected or not |
| `was replied to` | `boolean` | r/o | Indicates whether the message was replied to or not |

**Responds to:** `open`, `bounce`, `forward`, `redirect`, `reply`

### `account`

A Mail account for receiving messages (POP/IMAP). To create a new receiving account, use the 'pop account', 'imap account', and 'iCloud account' objects

**Contains:** `mailbox`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `delivery account` | `smtp server | missing value` | r/w | The delivery account used when sending mail from this account |
| `name` | `text` | r/w | The name of an account |
| `id` | `text` | r/o | The unique identifier of the account |
| `password` | `text` | r/w | Password for this account. Can be set, but not read via scripting |
| `authentication` | `Authentication` | r/w | Preferred authentication scheme for account |
| `account type` | `TypeOfAccount` | r/o | The type of an account |
| `email addresses` | `text` | r/w | The list of email addresses configured for an account |
| `full name` | `text` | r/w | The users full name configured for an account |
| `empty junk messages frequency` | `integer` | r/w | Number of days before junk messages are deleted (0 = delete on quit, -1 = never delete) |
| `empty sent messages frequency` | `integer` | r/w | Does nothing at all (deprecated) |
| `empty trash frequency` | `integer` | r/w | Number of days before messages in the trash are permanently deleted (0 = delete on quit, -1 = never delete) |
| `empty junk messages on quit` | `boolean` | r/w | Indicates whether the messages in the junk messages mailboxes will be deleted on quit |
| `empty sent messages on quit` | `boolean` | r/w | Does nothing at all (deprecated) |
| `empty trash on quit` | `boolean` | r/w | Indicates whether the messages in deleted messages mailboxes will be permanently deleted on quit |
| `enabled` | `boolean` | r/w | Indicates whether the account is enabled or not |
| `user name` | `text` | r/w | The user name used to connect to an account |
| `account directory` | `file` | r/o | The directory where the account stores things on disk |
| `port` | `integer` | r/w | The port used to connect to an account |
| `server name` | `text` | r/w | The host name used to connect to an account |
| `include when getting new mail` | `boolean` | r/w | Does nothing at all (deprecated) |
| `move deleted messages to trash` | `boolean` | r/w | Indicates whether messages that are deleted will be moved to the trash mailbox |
| `uses ssl` | `boolean` | r/w | Indicates whether SSL is enabled for this receiving account |

### `imap account`

An IMAP email account

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `compact mailboxes when closing` | `boolean` | r/w | Indicates whether an IMAP mailbox is automatically compacted when you quit Mail or switch to another mailbox |
| `message caching` | `MessageCachingPolicy` | r/w | Message caching setting for this account |
| `store drafts on server` | `boolean` | r/w | Indicates whether drafts will be stored on the IMAP server |
| `store junk mail on server` | `boolean` | r/w | Indicates whether junk mail will be stored on the IMAP server |
| `store sent messages on server` | `boolean` | r/w | Indicates whether sent messages will be stored on the IMAP server |
| `store deleted messages on server` | `boolean` | r/w | Indicates whether deleted messages will be stored on the IMAP server |

### `iCloud account`

An iCloud or MobileMe email account

### `pop account`

A POP email account

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `big message warning size` | `integer` | r/w | If message size (in bytes) is over this amount, Mail will prompt you asking whether you want to download the message (-1 = do not prompt) |
| `delayed message deletion interval` | `integer` | r/w | Number of days before messages that have been downloaded will be deleted from the server (0 = delete immediately after downloading) |
| `delete mail on server` | `boolean` | r/w | Indicates whether POP account deletes messages on the server after downloading |
| `delete messages when moved from inbox` | `boolean` | r/w | Indicates whether messages will be deleted from the server when moved from your POP inbox |

### `smtp server`

An SMTP account (for sending email)

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of an account |
| `password` | `text` | r/w | Password for this account. Can be set, but not read via scripting |
| `account type` | `TypeOfAccount` | r/o | The type of an account |
| `authentication` | `Authentication` | r/w | Preferred authentication scheme for account |
| `enabled` | `boolean` | r/w | Indicates whether the account is enabled or not |
| `user name` | `text` | r/w | The user name used to connect to an account |
| `port` | `integer` | r/w | The port used to connect to an account |
| `server name` | `text` | r/w | The host name used to connect to an account |
| `uses ssl` | `boolean` | r/w | Indicates whether SSL is enabled for this receiving account |

### `mailbox`

A mailbox that holds messages

*Plural:* `mailboxes`

**Contains:** `mailbox`, `message`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/w | The name of a mailbox |
| `unread count` | `integer` | r/o | The number of unread messages in the mailbox |
| `account` | `account` | r/o |  |
| `container` | `mailbox` | r/o |  |

### `rule`

Class for message rules

**Contains:** `rule condition`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color message` | `HighlightColors` | r/w | If rule matches, apply this color |
| `delete message` | `boolean` | r/w | If rule matches, delete message |
| `forward text` | `text` | r/w | If rule matches, prepend this text to the forwarded message. Set to empty string to include no prepended text |
| `forward message` | `text` | r/w | If rule matches, forward message to this address, or multiple addresses, separated by commas. Set to empty string to disable this action |
| `mark flagged` | `boolean` | r/w | If rule matches, mark message as flagged |
| `mark flag index` | `integer` | r/w | If rule matches, mark message with the specified flag. Set to -1 to disable this action |
| `mark read` | `boolean` | r/w | If rule matches, mark message as read |
| `play sound` | `text` | r/w | If rule matches, play this sound (specify name of sound or path to sound) |
| `redirect message` | `text` | r/w | If rule matches, redirect message to this address or multiple addresses, separate by commas. Set to empty string to disable this action |
| `reply text` | `text` | r/w | If rule matches, reply to message and prepend with this text. Set to empty string to disable this action |
| `run script` | `file | missing value` | r/w | If rule matches, run this compiled AppleScript file. Set to empty string to disable this action |
| `all conditions must be met` | `boolean` | r/w | Indicates whether all conditions must be met for rule to execute |
| `copy message` | `mailbox` | r/w | If rule matches, copy to this mailbox |
| `move message` | `mailbox` | r/w | If rule matches, move to this mailbox |
| `highlight text using color` | `boolean` | r/w | Indicates whether the color will be used to highlight the text or background of a message in the message list |
| `enabled` | `boolean` | r/w | Indicates whether the rule is enabled |
| `name` | `text` | r/w | Name of rule |
| `should copy message` | `boolean` | r/w | Indicates whether the rule has a copy action |
| `should move message` | `boolean` | r/w | Indicates whether the rule has a move action |
| `stop evaluating rules` | `boolean` | r/w | If rule matches, stop rule evaluation for this message |

### `rule condition`

Class for conditions that can be attached to a single rule

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `expression` | `text` | r/w | Rule expression field |
| `header` | `text` | r/w | Rule header key |
| `qualifier` | `RuleQualifier` | r/w | Rule qualifier |
| `rule type` | `RuleType` | r/w | Rule type |

### `recipient`

An email recipient

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `address` | `text` | r/w | The recipients email address |
| `name` | `text` | r/w | The name used for display |

### `bcc recipient`

An email recipient in the Bcc: field

### `cc recipient`

An email recipient in the Cc: field

### `to recipient`

An email recipient in the To: field

### `container`

A mailbox that contains other mailboxes.

### `header`

A header value for a message. E.g. To, Subject, From.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `content` | `text` | r/w | Contents of the header |
| `name` | `text` | r/w | Name of the header value |

### `mail attachment`

A file attached to a received message.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | Name of the attachment |
| `MIME type` | `text` | r/o | MIME type of the attachment E.g. text/plain. |
| `file size` | `integer` | r/o | Approximate size in bytes. |
| `downloaded` | `boolean` | r/o | Indicates whether the attachment has been downloaded. |
| `id` | `text` | r/o | The unique identifier of the attachment. |

**Responds to:** `save`

### Enumerations

### `Authentication`

| Value | Description |
|-------|-------------|
| `password` | Clear text password |
| `apop` | APOP |
| `kerberos 5` | Kerberos V5 (GSSAPI) |
| `ntlm` | NTLM |
| `md5` | CRAM-MD5 |
| `external` | External authentication (TLS client certificate) |
| `Apple token` | Apple token |
| `none` | None |

### `HighlightColors`

| Value | Description |
|-------|-------------|
| `blue` | Blue |
| `gray` | Gray |
| `green` | Green |
| `none` | None |
| `orange` | Orange |
| `other` | Other |
| `purple` | Purple |
| `red` | Red |
| `yellow` | Yellow |

### `MessageCachingPolicy`

| Value | Description |
|-------|-------------|
| `do not keep copies of any messages` | Do not use this option (deprecated). If you do, Mail will use the 'all messages but omit attachments' policy |
| `only messages I have read` | Do not use this option (deprecated). If you do, Mail will use the 'all messages but omit attachments' policy |
| `all messages but omit attachments` | All messages but omit attachments |
| `all messages and their attachments` | All messages and their attachments |

### `RuleQualifier`

| Value | Description |
|-------|-------------|
| `begins with value` | Begins with value |
| `does contain value` | Does contain value |
| `does not contain value` | Does not contain value |
| `ends with value` | Ends with value |
| `equal to value` | Equal to value |
| `less than value` | Less than value |
| `greater than value` | Greater than value |
| `none` | Indicates no qualifier is applicable |

### `RuleType`

| Value | Description |
|-------|-------------|
| `account` | Account |
| `any recipient` | Any recipient |
| `cc header` | Cc header |
| `matches every message` | Every message |
| `from header` | From header |
| `header key` | An arbitrary header key |
| `message content` | Message content |
| `message is junk mail` | Message is junk mail |
| `sender is in my contacts` | Sender is in my contacts |
| `sender is in my previous recipients` | Sender is in my previous recipients |
| `sender is member of group` | Sender is member of group |
| `sender is not in my contacts` | Sender is not in my contacts |
| `sender is not in my previous recipients` | sender is not in my previous recipients |
| `sender is not member of group` | Sender is not member of group |
| `sender is VIP` | Sender is VIP |
| `subject header` | Subject header |
| `to header` | To header |
| `to or cc header` | To or Cc header |
| `attachment type` | Attachment Type |

### `TypeOfAccount`

| Value | Description |
|-------|-------------|
| `pop` | POP |
| `smtp` | SMTP |
| `imap` | IMAP |
| `iCloud` | iCloud |
| `unknown` | Unknown |
