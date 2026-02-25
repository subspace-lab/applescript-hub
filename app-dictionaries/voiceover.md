# Voiceover AppleScript Dictionary

> Auto-generated from `VoiceOver.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Voiceover"`

## Table of Contents

- [VoiceOver Suite](#voiceover-suite)

---

## VoiceOver Suite

VOASeOver AppleScripting facilities

### Commands

### `perform command`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | The English name of the VoiceOver command to perform |

### `grab screenshot`

Takes a screenshot of the VO cursor and returns the path to the file

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `vo cursor object` | no | the vo cursor object to grab |

**Returns:** `text` — the path to the screenshot

### `click`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `click count` | yes | Number of times to click |
| `with` | `click button` | yes | Mouse button to click |

### `quit`

### `press`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `mouse cursor object` | no | The mouse cursor object |

### `release`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `mouse cursor object` | no | The mouse cursor object |

### `perform action`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `vo cursor object` | no | The vo cursor object |

### `select`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `vo cursor object` | no | The vo cursor object |

### `move`

Move the vo cursor to a new location.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `direction | containment` | yes | The direction to move in. |
| `to` | `place` | yes | The place to move to |

### `output`

Output

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text | outputables` | yes | item to be outputted |
| `with` | `spelling type` | yes | Type of reading |

### `open`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `menu | resource` | no | the item to open |

### `close menu`

Closes open menus

### `save`

Save last phrase

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `last phrase object` | no | The last phrase object |

### `copy to pasteboard`

Copy last phrase to pasteboard

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `last phrase object` | no | The last phrase object |

### Classes

### `application`

VoiceOver

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `vo cursor` | `vo cursor object` | r/o | The VoiceOver cursor |
| `commander` | `commander object` | r/o | The VoiceOver commander |
| `mouse cursor` | `mouse cursor object` | r/o | The mouse cursor |
| `keyboard cursor` | `keyboard cursor object` | r/o | The keyboard cursor |
| `caption window` | `caption window object` | r/o | The VoiceOver caption window |
| `braille window` | `braille window object` | r/o | The VoiceOver Braille window |
| `last phrase` | `last phrase object` | r/o | The last phrase VoiceOver output |

**Responds to:** `output`, `open`, `close menu`, `quit`

### `vo cursor object`

The VoiceOver cursor

*Plural:* `vo cursor object`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bounds` | `rectangle` | r/o | The bounds of the VoiceOver cursor |
| `text under cursor` | `text` | r/o | The text of the item in the VoiceOver cursor |
| `magnification` | `real` | r/w | The magnification factor of the VoiceOver cursor |

**Responds to:** `grab screenshot`, `move`, `perform action`, `select`

### `commander object`

The VoiceOver commander.

**Responds to:** `perform command`

### `caption window object`

The caption window

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `enabled` | `boolean` | r/w | The visibility of the caption window |

### `braille window object`

The Braille window

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `enabled` | `boolean` | r/w | The visibility of the Braille window |

### `mouse cursor object`

The mouse cursor

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `position` | `point` | r/w | Position of the mouse |

**Responds to:** `click`, `press`, `release`

### `last phrase object`

The last phrase outputted

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `content` | `text` | r/o | The text of the last phrase |

**Responds to:** `save`, `copy to pasteboard`, `output`

### `keyboard cursor object`

The keyboard cursor

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bounds` | `rectangle` | r/o | The bounds of the keyboard cursor |
| `text under cursor` | `text` | r/o | The text of the item in the keyboard cursor |

### Enumerations

### `click count`

| Value | Description |
|-------|-------------|
| `once` | One mouse click |
| `twice` | Double mouse click |
| `thrice` | Triple mouse click |

### `click button`

| Value | Description |
|-------|-------------|
| `left button` | Left mouse button |
| `right button` | Right mouse button |

### `place`

| Value | Description |
|-------|-------------|
| `dock` | The dock |
| `desktop` | The desktop |
| `menubar` | The menubar |
| `status menu` | Status menu |
| `spotlight` | Spotlight |
| `linked item` | Linked item |
| `first item` | First item |
| `last item` | Last item |

### `direction`

| Value | Description |
|-------|-------------|
| `up` | Up |
| `down` | Down |
| `left` | Left |
| `right` | Right |

### `containment`

| Value | Description |
|-------|-------------|
| `into item` | Interact in |
| `out of item` | Interact out |

### `menu`

| Value | Description |
|-------|-------------|
| `help menu` | Help menu |
| `applications menu` | Applications menu |
| `windows menu` | Windows menu |
| `commands menu` | Commands menu |
| `item chooser` | Item chooser |
| `web menu` | Web menu |
| `contextual menu` | Contextual menu |

### `resource`

| Value | Description |
|-------|-------------|
| `utility` | VoiceOver Utility |
| `quickstart` | Quickstart |
| `VoiceOver help` | Quickstart |

### `outputables`

| Value | Description |
|-------|-------------|
| `mouse summary` | Summary of the item under the mouse |
| `workspace overview` | The overview of the working environment |
| `window overview` | The overview of the current window |
| `web overview` | The overview of the web page |
| `announcement history` | Causes the display to show recent announcement |

### `spelling type`

| Value | Description |
|-------|-------------|
| `alphabetic spelling` | Alphabetic spelling |
| `phonetic spelling` | Phoenetic spelling |
