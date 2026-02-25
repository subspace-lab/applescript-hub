# Systemevents AppleScript Dictionary

> Auto-generated from `SystemEvents.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Systemevents"`

## Table of Contents

- [System Events Suite](#system-events-suite)
- [Accounts Suite](#accounts-suite)
- [Appearance Suite](#appearance-suite)
- [CD and DVD Preferences Suite](#cd-and-dvd-preferences-suite)
- [Desktop Suite](#desktop-suite)
- [Dock Preferences Suite](#dock-preferences-suite)
- [Login Items Suite](#login-items-suite)
- [Network Preferences Suite](#network-preferences-suite)
- [Screen Saver Suite](#screen-saver-suite)
- [Security Suite](#security-suite)
- [Disk-Folder-File Suite](#disk-folder-file-suite)
- [Power Suite](#power-suite)
- [Processes Suite](#processes-suite)
- [Property List Suite](#property-list-suite)
- [XML Suite](#xml-suite)
- [Type Definitions](#type-definitions)
- [Hidden Suite](#hidden-suite)
- [Scripting Definition Suite](#scripting-definition-suite)

---

## System Events Suite

Terms and Events for controlling the System Events application

### Commands

### `abort transaction`

Discard the results of a bounded update session with one or more files.

### `begin transaction`

Begin a bounded update session with one or more files.

**Returns:** `integer`

### `end transaction`

Apply the results of a bounded update session with one or more files.

### Enumerations

### `saveable file format`

| Value | Description |
|-------|-------------|
| `text` | Text File Format |

---

## Accounts Suite

Terms and Events for controlling the users account settings

### Classes

### `user`

user account

*Plural:* `users`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `full name` | `text` | r/o | user's full name |
| `home directory` | `text | file` | r/o | path to user's home directory |
| `name` | `text` | r/o | user's short name |
| `picture path` | `text | file` | r/w | path to user's picture. Can be set for current user only! |

---

## Appearance Suite

Terms for controlling Appearance preferences

### Classes

### `appearance preferences object`

A collection of appearance preferences

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `appearance` | `Appearances` | r/w | the overall look of buttons, menus and windows |
| `font smoothing` | `boolean` | r/w | Is font smoothing on? |
| `font smoothing limit` | `integer` | r/o | the font size at or below which font smoothing is turned off |
| `font smoothing style` | `FontSmoothingStyles` | r/w | the method used for smoothing fonts |
| `highlight color` | `HighlightColors | color` | r/w | color used for hightlighting selected text and lists |
| `recent applications limit` | `integer` | r/w | the number of recent applications to track |
| `recent documents limit` | `integer` | r/w | the number of recent documents to track |
| `recent servers limit` | `integer` | r/w | the number of recent servers to track |
| `scroll bar action` | `ScrollPageBehaviors` | r/w | the action performed by clicking the scroll bar |
| `smooth scrolling` | `boolean` | r/w | Is smooth scrolling used? |
| `dark mode` | `boolean` | r/w | use dark menu bar and dock |

### Enumerations

### `ScrollPageBehaviors`

| Value | Description |
|-------|-------------|
| `jump to here` | jump to here |
| `jump to next page` | jump to next page |

### `FontSmoothingStyles`

| Value | Description |
|-------|-------------|
| `automatic` | automatic |
| `light` | light |
| `medium` | medium |
| `standard` | standard |
| `strong` | strong |

### `Appearances`

| Value | Description |
|-------|-------------|
| `blue` | blue |
| `graphite` | graphite |

### `HighlightColors`

| Value | Description |
|-------|-------------|
| `blue` | blue |
| `gold` | gold |
| `graphite` | graphite |
| `green` | green |
| `orange` | orange |
| `purple` | purple |
| `red` | red |
| `silver` | silver |

---

## CD and DVD Preferences Suite

Terms and Events for controlling the actions when inserting CDs and DVDs

### Classes

### `CD and DVD preferences object`

user's CD and DVD insertion preferences

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `blank CD` | `insertion preference` | r/o | the blank CD insertion preference |
| `blank DVD` | `insertion preference` | r/o | the blank DVD insertion preference |
| `blank BD` | `insertion preference` | r/o | the blank BD insertion preference |
| `music CD` | `insertion preference` | r/o | the music CD insertion preference |
| `picture CD` | `insertion preference` | r/o | the picture CD insertion preference |
| `video DVD` | `insertion preference` | r/o | the video DVD insertion preference |
| `video BD` | `insertion preference` | r/o | the video BD insertion preference |

### `insertion preference`

a specific insertion preference

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `custom application` | `text | missing value` | r/w | application to launch or activate on the insertion of media |
| `custom script` | `text | missing value` | r/w | AppleScript to launch or activate on the insertion of media |
| `insertion action` | `dhac` | r/w | action to perform on media insertion |

### Enumerations

### `dhac`

| Value | Description |
|-------|-------------|
| `ask what to do` | ask what to do |
| `ignore` | ignore |
| `open application` | open application |
| `run a script` | run a script |

---

## Desktop Suite

Terms and Events for controlling the desktop picture settings.

### Classes

### `desktop`

desktop picture settings

*Plural:* `desktops`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | name of the desktop |
| `id` | `integer` | r/o | unique identifier of the desktop |
| `change interval` | `real` | r/w | number of seconds to wait between changing the desktop picture |
| `display name` | `text` | r/o | name of display on which this desktop appears |
| `picture` | `text | file` | r/w | path to file used as desktop picture |
| `picture rotation` | `integer` | r/w | never, using interval, using login, after sleep |
| `pictures folder` | `text | file` | r/w | path to folder containing pictures for changing desktop background |
| `random order` | `boolean` | r/w | turn on for random ordering of changing desktop pictures |
| `translucent menu bar` | `boolean` | r/w | indicates whether the menu bar is translucent |
| `dynamic style` | `dynamic style` | r/w | desktop picture dynamic style |

### Enumerations

### `dynamic style`

| Value | Description |
|-------|-------------|
| `auto` | automatic (if supported, follows light/dark appearance) |
| `dynamic` | dynamic (if supported, updates desktop picture based on time and/or location) |
| `light` | light |
| `dark` | dark |
| `unknown` | unknown value |

---

## Dock Preferences Suite

Terms and Events for controlling the dock preferences

### Classes

### `dock preferences object`

user's dock preferences

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `animate` | `boolean` | r/w | is the animation of opening applications on or off? |
| `autohide` | `boolean` | r/w | is autohiding the dock on or off? |
| `dock size` | `real` | r/w | size/height of the items (between 0.0 (minimum) and 1.0 (maximum)) |
| `autohide menu bar` | `boolean` | r/w | is autohiding the menu bar on or off? |
| `double click behavior` | `dpbh` | r/w | behaviour when double clicking window a title bar |
| `magnification` | `boolean` | r/w | is magnification on or off? |
| `magnification size` | `real` | r/w | maximum magnification size when magnification is on (between 0.0 (minimum) and 1.0 (maximum)) |
| `minimize effect` | `dpef` | r/w | minimization effect |
| `minimize into application` | `boolean` | r/w | minimize window into its application? |
| `screen edge` | `dpls` | r/w | location on screen |
| `show indicators` | `boolean` | r/w | show indicators for open applications? |
| `show recents` | `boolean` | r/w | show recent applications? |

### Enumerations

### `dpls`

| Value | Description |
|-------|-------------|
| `bottom` | bottom |
| `left` | left |
| `right` | right |

### `dpef`

| Value | Description |
|-------|-------------|
| `genie` | genie |
| `scale` | scale |

### `dpbh`

| Value | Description |
|-------|-------------|
| `minimize` | minimize |
| `off` | off |
| `zoom` | zoom |

---

## Login Items Suite

Terms and Events for controlling the Login Items application

### Classes

### `login item`

an item to be launched or opened at login

*Plural:* `login items`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `hidden` | `boolean` | r/w | Is the Login Item hidden when launched? |
| `kind` | `text` | r/o | the file type of the Login Item |
| `name` | `text` | r/o | the name of the Login Item |
| `path` | `text` | r/o | the file system path to the Login Item |

---

## Network Preferences Suite

Terms and Commands for manipulating and viewing network settings

### Commands

### `connect`

connect a configuration or service

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `configuration | service` | no | a configuration or service |

**Returns:** `configuration`

### `disconnect`

disconnect a configuration or service

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `configuration | service` | no | a configuration or service |

**Returns:** `configuration`

### Classes

### `configuration`

A collection of settings for configuring a connection

*Plural:* `configurations`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `account name` | `text` | r/w | the name used to authenticate |
| `connected` | `boolean` | r/o | Is the configuration connected? |
| `id` | `text` | r/o | the unique identifier for the configuration |
| `name` | `text` | r/o | the name of the configuration |

**Responds to:** `connect`, `disconnect`

### `interface`

A collection of settings for a network interface

*Plural:* `interfaces`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `automatic` | `boolean` | r/w | configure the interface speed, duplex, and mtu automatically? |
| `duplex` | `text` | r/w | the duplex setting  half | full | full with flow control |
| `id` | `text` | r/o | the unique identifier for the interface |
| `kind` | `text` | r/o | the type of interface |
| `MAC address` | `text` | r/o | the MAC address for the interface |
| `mtu` | `integer` | r/w | the packet size |
| `name` | `text` | r/o | the name of the interface |
| `speed` | `integer` | r/w | ethernet speed 10 | 100 | 1000 |

### `location`

A set of services

*Plural:* `locations`

**Contains:** `service`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | the unique identifier for the location |
| `name` | `text` | r/w | the name of the location |

### `network preferences object`

the preferences for the current user's network

**Contains:** `interface`, `location`, `service`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `current location` | `location` | r/w | the current location |

### `service`

A collection of settings for a network service

*Plural:* `services`

**Contains:** `configuration`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `active` | `boolean` | r/o | Is the service active? |
| `current configuration` | `configuration` | r/w | the currently selected configuration |
| `id` | `text` | r/o | the unique identifier for the service |
| `interface` | `interface` | r/o | the interface the service is built on |
| `kind` | `integer` | r/o | the type of service |
| `name` | `text` | r/w | the name of the service |

**Responds to:** `connect`, `disconnect`

---

## Screen Saver Suite

Terms and Events for controlling screen saver settings.

### Commands

### `start`

start the screen saver

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `screen saver | screen saver preferences object` | no | the object for the command |

### `stop`

stop the screen saver

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `screen saver | screen saver preferences object` | no | the object for the command |

### Classes

### `screen saver`

an installed screen saver

*Plural:* `screen savers`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `displayed name` | `text` | r/o | name of the screen saver module as displayed to the user |
| `name` | `text` | r/o | name of the screen saver module to be displayed |
| `path` | `alias` | r/o | path to the screen saver module |
| `picture display style` | `text` | r/w | effect to use when displaying picture-based screen savers (slideshow, collage, or mosaic) |

**Responds to:** `start`, `stop`

### `screen saver preferences object`

screen saver settings

*Plural:* `screen saver preferences objects`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `delay interval` | `integer` | r/w | number of seconds of idle time before the screen saver starts; zero for never |
| `main screen only` | `boolean` | r/w | should the screen saver be shown only on the main screen? |
| `running` | `boolean` | r/o | is the screen saver running? |
| `show clock` | `boolean` | r/w | should a clock appear over the screen saver? |

**Responds to:** `start`, `stop`

---

## Security Suite

Terms for controlling Security preferences

### Classes

### `security preferences object`

a collection of security preferences

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `automatic login` | `boolean` | r/w | Is automatic login allowed? |
| `log out when inactive` | `boolean` | r/w | Will the computer log out when inactive? |
| `log out when inactive interval` | `integer` | r/w | The interval of inactivity after which the computer will log out |
| `require password to unlock` | `boolean` | r/w | Is a password required to unlock secure preferences? |
| `require password to wake` | `boolean` | r/w | Is a password required to wake the computer from sleep or screen saver? |
| `secure virtual memory` | `boolean` | r/w | Is secure virtual memory being used? |

---

## Disk-Folder-File Suite

Terms and Events for controlling Disks, Folders, and Files

### Commands

### `delete`

Delete disk item(s).

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `disk item` | no | The disk item(s) to be deleted. |

### `move`

Move disk item(s) to a new location.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `disk item | disk item | text | text` | no | The disk item(s) to be moved. |
| `to` | `location specifier | text` | no | The new location for the disk item(s). |

**Returns:** `disk item | disk item`

### `open`

Open disk item(s) with the appropriate application.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `file | text` | no | The disk item(s) to be opened. |

**Returns:** `file`

### Classes

### `alias`

An alias in the file system

*Plural:* `aliases`

**Contains:** `alias`, `disk item`, `file`, `file package`, `folder`, `item`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `creator type` | `text | any` | r/w | the OSType identifying the application that created the alias |
| `default application` | `disk item | any` | r/w | the application that will launch if the alias is opened |
| `file type` | `text | any` | r/w | the OSType identifying the type of data contained in the alias |
| `kind` | `text` | r/o | The kind of alias, as shown in Finder |
| `product version` | `text` | r/o | the version of the product (visible at the top of the "Get Info" window) |
| `short version` | `text` | r/o | the short version of the application bundle referenced by the alias |
| `stationery` | `boolean` | r/w | Is the alias a stationery pad? |
| `type identifier` | `text` | r/o | The type identifier of the alias |
| `version` | `text` | r/o | the version of the application bundle referenced by the alias (visible at the bottom of the "Get Info" window) |

**Responds to:** `delete`

### `Classic domain object`

The Classic domain in the file system

*Plural:* `Classic domain objects`

**Contains:** `folder`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `apple menu folder` | `folder` | r/o | The Apple Menu Items folder |
| `control panels folder` | `folder` | r/o | The Control Panels folder |
| `control strip modules folder` | `folder` | r/o | The Control Strip Modules folder |
| `desktop folder` | `folder` | r/o | The Classic Desktop folder |
| `extensions folder` | `folder` | r/o | The Extensions folder |
| `fonts folder` | `folder` | r/o | The Fonts folder |
| `launcher items folder` | `folder` | r/o | The Launcher Items folder |
| `preferences folder` | `folder` | r/o | The Classic Preferences folder |
| `shutdown folder` | `folder` | r/o | The Shutdown Items folder |
| `startup items folder` | `folder` | r/o | The StartupItems folder |
| `system folder` | `folder` | r/o | The System folder |

### `disk`

A disk in the file system

*Plural:* `disks`

**Contains:** `alias`, `disk item`, `file`, `file package`, `folder`, `item`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `capacity` | `number` | r/o | the total number of bytes (free or used) on the disk |
| `ejectable` | `boolean` | r/o | Can the media be ejected (floppies, CD's, and so on)? |
| `format` | `edfm` | r/o | the file system format of this disk |
| `free space` | `number` | r/o | the number of free bytes left on the disk |
| `ignore privileges` | `boolean` | r/w | Ignore permissions on this disk? |
| `local volume` | `boolean` | r/o | Is the media a local volume (as opposed to a file server)? |
| `server` | `text | any` | r/o | the server on which the disk resides, AFP volumes only |
| `startup` | `boolean` | r/o | Is this disk the boot disk? |
| `zone` | `text | any` | r/o | the zone in which the disk's server resides, AFP volumes only |

**Responds to:** `delete`

### `disk item`

An item stored in the file system

*Plural:* `disk items`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `busy status` | `boolean` | r/o | Is the disk item busy? |
| `container` | `disk item` | r/o | the folder or disk which has this disk item as an element |
| `creation date` | `date` | r/o | the date on which the disk item was created |
| `displayed name` | `text` | r/o | the name of the disk item as displayed in the User Interface |
| `id` | `text` | r/o | the unique ID of the disk item |
| `modification date` | `date` | r/w | the date on which the disk item was last modified |
| `name` | `text` | r/w | the name of the disk item |
| `name extension` | `text` | r/o | the extension portion of the name |
| `package folder` | `boolean` | r/o | Is the disk item a package? |
| `path` | `text` | r/o | the file system path of the disk item |
| `physical size` | `integer` | r/o | the actual space used by the disk item on disk |
| `POSIX path` | `text` | r/o | the POSIX file system path of the disk item |
| `size` | `integer` | r/o | the logical size of the disk item |
| `URL` | `text` | r/o | the URL of the disk item |
| `visible` | `boolean` | r/w | Is the disk item visible? |
| `volume` | `text` | r/o | the volume on which the disk item resides |

**Responds to:** `open`, `move`

### `domain`

A domain in the file system

*Plural:* `domains`

**Contains:** `folder`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `application support folder` | `folder` | r/o | The Application Support folder |
| `applications folder` | `folder` | r/o | The Applications folder |
| `desktop pictures folder` | `folder` | r/o | The Desktop Pictures folder |
| `Folder Action scripts folder` | `folder` | r/o | The Folder Action Scripts folder |
| `fonts folder` | `folder` | r/o | The Fonts folder |
| `id` | `text` | r/o | the unique identifier of the domain |
| `library folder` | `folder` | r/o | The Library folder |
| `name` | `text` | r/o | the name of the domain |
| `preferences folder` | `folder` | r/o | The Preferences folder |
| `scripting additions folder` | `folder` | r/o | The Scripting Additions folder |
| `scripts folder` | `folder` | r/o | The Scripts folder |
| `shared documents folder` | `folder` | r/o | The Shared Documents folder |
| `speakable items folder` | `folder` | r/o | The Speakable Items folder |
| `utilities folder` | `folder` | r/o | The Utilities folder |
| `workflows folder` | `folder` | r/o | The Automator Workflows folder |

### `file`

A file in the file system

*Plural:* `files`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `creator type` | `text | any` | r/w | the OSType identifying the application that created the file |
| `default application` | `disk item | any` | r/w | the application that will launch if the file is opened |
| `file type` | `text | any` | r/w | the OSType identifying the type of data contained in the file |
| `kind` | `text` | r/o | The kind of file, as shown in Finder |
| `product version` | `text` | r/o | the version of the product (visible at the top of the "Get Info" window) |
| `short version` | `text` | r/o | the short version of the file |
| `stationery` | `boolean` | r/w | Is the file a stationery pad? |
| `type identifier` | `text` | r/o | The type identifier of the file |
| `version` | `text` | r/o | the version of the file (visible at the bottom of the "Get Info" window) |

### `file package`

A file package in the file system

*Plural:* `file packages`

**Contains:** `alias`, `disk item`, `file`, `file package`, `folder`, `item`

**Responds to:** `delete`

### `folder`

A folder in the file system

*Plural:* `folders`

**Contains:** `alias`, `disk item`, `file`, `file package`, `folder`, `item`

**Responds to:** `delete`

### `local domain object`

The local domain in the file system

*Plural:* `local domain objects`

**Contains:** `folder`

### `network domain object`

The network domain in the file system

*Plural:* `network domain objects`

**Contains:** `folder`

### `system domain object`

The system domain in the file system

*Plural:* `system domain objects`

**Contains:** `folder`

### `user domain object`

The user domain in the file system

*Plural:* `user domain objects`

**Contains:** `folder`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `desktop folder` | `folder` | r/o | The user's Desktop folder |
| `documents folder` | `folder` | r/o | The user's Documents folder |
| `downloads folder` | `folder` | r/o | The user's Downloads folder |
| `favorites folder` | `folder` | r/o | The user's Favorites folder |
| `home folder` | `folder` | r/o | The user's Home folder |
| `movies folder` | `folder` | r/o | The user's Movies folder |
| `music folder` | `folder` | r/o | The user's Music folder |
| `pictures folder` | `folder` | r/o | The user's Pictures folder |
| `public folder` | `folder` | r/o | The user's Public folder |
| `sites folder` | `folder` | r/o | The user's Sites folder |
| `temporary items folder` | `folder` | r/o | The Temporary Items folder |

### Enumerations

### `edfm`

| Value | Description |
|-------|-------------|
| `Apple Photo format` | Apple Photo format |
| `AppleShare format` | AppleShare format |
| `audio format` | audio format |
| `High Sierra format` | High Sierra format |
| `ISO 9660 format` | ISO 9660 format |
| `Mac OS Extended format` | Mac OS Extended format |
| `Mac OS format` | Mac OS format |
| `MSDOS format` | MSDOS format |
| `NFS format` | NFS format |
| `ProDOS format` | ProDOS format |
| `QuickTake format` | QuickTake format |
| `UDF format` | UDF format |
| `UFS format` | UFS format |
| `unknown format` | unknown format |
| `WebDAV format` | WebDAV format |

---

## Power Suite

Terms and Events for controlling System power

### Commands

### `log out`

Log out the current user

### `restart`

Restart the computer

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `state saving preference` | `boolean` | yes | Is the user defined state saving preference followed? |

### `shut down`

Shut Down the computer

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `state saving preference` | `boolean` | yes | Is the user defined state saving preference followed? |

### `sleep`

Put the computer to sleep

---

## Processes Suite

Terms and Events for controlling Processes

### Commands

### `click`

cause the target process to behave as if the UI element were clicked

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `UI element` | yes | The UI element to be clicked. |
| `at` | `number` | yes | when sent to a "process" object, the { x, y } location at which to click, in global coordinates |

**Returns:** `UI element | UI element`

### `key code`

cause the target process to behave as if key codes were entered

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `integer | integer` | no | The key code(s) to be sent. May be a list. |
| `using` | `eMds | eMds` | yes | modifiers with which the key codes are to be entered |

### `keystroke`

cause the target process to behave as if keystrokes were entered

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | The keystrokes to be sent. |
| `using` | `eMds | eMds` | yes | modifiers with which the keystrokes are to be entered |

### `perform`

cause the target process to behave as if the action were applied to its UI element

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `action` | no | The action to be performed. |

**Returns:** `action`

### `select`

set the selected property of the UI element

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `UI element` | no | The UI element to be selected. |

**Returns:** `UI element`

### Classes

### `action`

An action that can be performed on the UI element

*Plural:* `actions`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `description` | `text` | r/o | what the action does |
| `name` | `text` | r/o | the name of the action |

**Responds to:** `perform`

### `application process`

A process launched from an application file

*Plural:* `application processes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `application file` | `alias | file | any` | r/o | a reference to the application file from which this process was launched |

### `attribute`

An named data value associated with the UI element

*Plural:* `attributes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | the name of the attribute |
| `settable` | `boolean` | r/o | Can the attribute be set? |
| `value` | `specifier | number | text | list | record | any | boolean` | r/w | the current value of the attribute |

### `browser`

A browser belonging to a window

*Plural:* `browsers`

### `busy indicator`

A busy indicator belonging to a window

*Plural:* `busy indicators`

### `button`

A button belonging to a window or scroll bar

*Plural:* `buttons`

### `checkbox`

A checkbox belonging to a window

*Plural:* `checkboxes`

### `color well`

A color well belonging to a window

*Plural:* `color wells`

### `column`

A column belonging to a table

*Plural:* `columns`

### `combo box`

A combo box belonging to a window

*Plural:* `combo boxes`

### `desk accessory process`

A process launched from an desk accessory file

*Plural:* `desk accessory processes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `desk accessory file` | `alias` | r/o | a reference to the desk accessory file from which this process was launched |

### `drawer`

A drawer that may be extended from a window

*Plural:* `drawers`

### `group`

A group belonging to a window

*Plural:* `groups`

**Contains:** `checkbox`, `static text`

### `grow area`

A grow area belonging to a window

*Plural:* `grow areas`

### `image`

An image belonging to a static text field

*Plural:* `images`

### `incrementor`

A incrementor belonging to a window

*Plural:* `incrementors`

### `list`

A list belonging to a window

*Plural:* `lists`

### `menu`

A menu belonging to a menu bar item

*Plural:* `menus`

**Contains:** `menu item`

### `menu bar`

A menu bar belonging to a process

*Plural:* `menu bars`

**Contains:** `menu`, `menu bar item`

### `menu bar item`

A menu bar item belonging to a menu bar

*Plural:* `menu bar items`

**Contains:** `menu`

### `menu button`

A menu button belonging to a window

*Plural:* `menu buttons`

### `menu item`

A menu item belonging to a menu

*Plural:* `menu items`

**Contains:** `menu`

### `outline`

A outline belonging to a window

*Plural:* `outlines`

### `pop over`

A pop over belonging to a window

*Plural:* `pop overs`

### `pop up button`

A pop up button belonging to a window

*Plural:* `pop up buttons`

### `process`

A process running on this computer

*Plural:* `processes`

**Contains:** `menu bar`, `window`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `accepts high level events` | `boolean` | r/o | Is the process high-level event aware (accepts open application, open document, print document, and quit)? |
| `accepts remote events` | `boolean` | r/o | Does the process accept remote events? |
| `architecture` | `text` | r/o | the architecture in which the process is running |
| `background only` | `boolean` | r/o | Does the process run exclusively in the background? |
| `bundle identifier` | `text` | r/o | the bundle identifier of the process' application file |
| `Classic` | `boolean` | r/o | Is the process running in the Classic environment? |
| `creator type` | `text` | r/o | the OSType of the creator of the process (the signature) |
| `displayed name` | `text` | r/o | the name of the file from which the process was launched, as displayed in the User Interface |
| `file` | `alias | file | any` | r/o | the file from which the process was launched |
| `file type` | `text` | r/o | the OSType of the file type of the process |
| `frontmost` | `boolean` | r/w | Is the process the frontmost process |
| `has scripting terminology` | `boolean` | r/o | Does the process have a scripting terminology, i.e., can it be scripted? |
| `id` | `integer` | r/o | The unique identifier of the process |
| `name` | `text` | r/o | the name of the process |
| `partition space used` | `integer` | r/o | the number of bytes currently used in the process' partition |
| `short name` | `text | missing value | any` | r/o | the short name of the file from which the process was launched |
| `total partition size` | `integer` | r/o | the size of the partition with which the process was launched |
| `unix id` | `integer` | r/o | The Unix process identifier of a process running in the native environment, or -1 for a process running in the Classic environment |
| `visible` | `boolean | number` | r/w | Is the process' layer visible? |

**Responds to:** `click`

### `progress indicator`

A progress indicator belonging to a window

*Plural:* `progress indicators`

### `radio button`

A radio button belonging to a window

*Plural:* `radio buttons`

### `radio group`

A radio button group belonging to a window

*Plural:* `radio groups`

**Contains:** `radio button`

### `relevance indicator`

A relevance indicator belonging to a window

*Plural:* `relevance indicators`

### `row`

A row belonging to a table

*Plural:* `rows`

### `scroll area`

A scroll area belonging to a window

*Plural:* `scroll areas`

### `scroll bar`

A scroll bar belonging to a window

*Plural:* `scroll bars`

**Contains:** `button`, `value indicator`

### `sheet`

A sheet displayed over a window

*Plural:* `sheets`

### `slider`

A slider belonging to a window

*Plural:* `sliders`

### `splitter`

A splitter belonging to a window

*Plural:* `splitters`

### `splitter group`

A splitter group belonging to a window

*Plural:* `splitter groups`

### `static text`

A static text field belonging to a window

*Plural:* `static texts`

**Contains:** `image`

### `tab group`

A tab group belonging to a window

*Plural:* `tab groups`

### `table`

A table belonging to a window

*Plural:* `tables`

### `text area`

A text area belonging to a window

*Plural:* `text areas`

### `text field`

A text field belonging to a window

*Plural:* `text fields`

### `toolbar`

A toolbar belonging to a window

*Plural:* `toolbars`

### `UI element`

A piece of the user interface of a process

*Plural:* `UI elements`

**Contains:** `action`, `attribute`, `browser`, `busy indicator`, `button`, `checkbox`, `color well`, `column`, `combo box`, `drawer`, `group`, `grow area`, `image`, `incrementor`, `list`, `menu`, `menu bar`, `menu bar item`, `menu button`, `menu item`, `outline`, `pop over`, `pop up button`, `progress indicator`, `radio button`, `radio group`, `relevance indicator`, `row`, `scroll area`, `scroll bar`, `sheet`, `slider`, `splitter`, `splitter group`, `static text`, `tab group`, `table`, `text area`, `text field`, `toolbar`, `UI element`, `value indicator`, `window`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `accessibility description` | `text | missing value | any` | r/o | a more complete description of the UI element and its capabilities |
| `class` | `type` | r/o | the class of the UI Element, which identifies it function |
| `description` | `text | missing value | any` | r/o | the accessibility description, if available; otherwise, the role description |
| `enabled` | `boolean | missing value | any` | r/o | Is the UI element enabled? ( Does it accept clicks? ) |
| `entire contents` | `specifier` | r/o | a list of every UI element contained in this UI element and its child UI elements, to the limits of the tree |
| `focused` | `boolean | missing value | any` | r/w | Is the focus on this UI element? |
| `help` | `text | missing value | any` | r/o | an elaborate description of the UI element and its capabilities |
| `maximum value` | `number | missing value | any` | r/o | the maximum value that the UI element can take on |
| `minimum value` | `number | missing value | any` | r/o | the minimum value that the UI element can take on |
| `name` | `text` | r/o | the name of the UI Element, which identifies it within its container |
| `orientation` | `text | missing value | any` | r/o | the orientation of the UI element |
| `position` | `number | missing value | any` | r/w | the position of the UI element |
| `role` | `text` | r/o | an encoded description of the UI element and its capabilities |
| `role description` | `text` | r/o | a more complete description of the UI element's role |
| `selected` | `boolean | missing value | any` | r/w | Is the UI element selected? |
| `size` | `number | missing value | any` | r/w | the size of the UI element |
| `subrole` | `text | missing value | any` | r/o | an encoded description of the UI element and its capabilities |
| `title` | `text` | r/o | the title of the UI element as it appears on the screen |
| `value` | `specifier | number | text | any` | r/w | the current value of the UI element |

**Responds to:** `click`, `select`, `increment`, `decrement`, `confirm`, `pick`, `cancel`

### `value indicator`

A value indicator ( thumb or slider ) belonging to a scroll bar

*Plural:* `value indicators`

### Enumerations

### `eMds`

| Value | Description |
|-------|-------------|
| `command down` | command down |
| `control down` | control down |
| `option down` | option down |
| `shift down` | shift down |

### `eMky`

| Value | Description |
|-------|-------------|
| `command` | command |
| `control` | control |
| `option` | option |
| `shift` | shift |

---

## Property List Suite

Terms and Events for accessing the content of Property List files

### Classes

### `property list file`

A file containing data in Property List format

*Plural:* `property list files`

### `data`

A data blob

### `property list item`

A unit of data in Property List format

*Plural:* `property list items`

**Contains:** `property list item`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `kind` | `type` | r/o | the kind of data stored in the property list item: boolean/data/date/list/number/record/string |
| `name` | `text` | r/o | the name of the property list item ( if any ) |
| `text` | `text` | r/w | the text representation of the property list data |
| `value` | `any | number | boolean | date | list | record | text | data` | r/w | the value of the property list item |

---

## XML Suite

Terms and Events for accessing the content of XML files

### Classes

### `XML attribute`

A named value associated with a unit of data in XML format

*Plural:* `XML attributes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | the name of the XML attribute |
| `value` | `boolean | date | list | number | record | text | any` | r/w | the value of the XML attribute |

### `XML data`

Data in XML format

**Contains:** `XML element`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | the unique identifier of the XML data |
| `name` | `text` | r/w | the name of the XML data |
| `text` | `text` | r/w | the text representation of the XML data |

### `XML element`

A unit of data in XML format

*Plural:* `XML elements`

**Contains:** `XML attribute`, `XML element`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | the unique identifier of the XML element |
| `name` | `text` | r/o | the name of the XML element |
| `value` | `boolean | date | list | number | record | text | any` | r/w | the value of the XML element |

### `XML file`

A file containing data in XML format

*Plural:* `XML files`

---

## Type Definitions

Records used in scripting System Events

### Classes

### `print settings`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `copies` | `integer` | r/w | the number of copies of a document to be printed |
| `collating` | `boolean` | r/w | Should printed copies be collated? |
| `starting page` | `integer` | r/w | the first page of the document to be printed |
| `ending page` | `integer` | r/w | the last page of the document to be printed |
| `pages across` | `integer` | r/w | number of logical pages laid across a physical page |
| `pages down` | `integer` | r/w | number of logical pages laid out down a physical page |
| `requested print time` | `date` | r/w | the time at which the desktop printer should print the document |
| `error handling` | `enum` | r/w | how errors are handled |
| `fax number` | `text` | r/w | for fax number |
| `target printer` | `text` | r/w | for target printer |

### Enumerations

### `enum`

| Value | Description |
|-------|-------------|
| `standard` | Standard PostScript error handling |
| `detailed` | print a detailed report of PostScript errors |

---

## Hidden Suite

Hidden Terms and Events for controlling the System Events application

### Commands

### `attach action to`

Attach an action to a folder

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |
| `using` | `text` | no | a file containing the script to attach |

**Returns:** `folder action`

### `attached scripts`

List the actions attached to a folder

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |

**Returns:** `any`

### `cancel`

cause the target process to behave as if the UI element were cancelled

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |

**Returns:** `UI element`

### `confirm`

cause the target process to behave as if the UI element were confirmed

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |

**Returns:** `UI element`

### `decrement`

cause the target process to behave as if the UI element were decremented

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |

**Returns:** `UI element`

### `do folder action`

Send a folder action code to a folder action script

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |
| `folder action code` | `actn` | no | the folder action message to process |
| `with item list` | `any` | yes | a list of items for the folder action message to process |
| `with window size` | `rectangle` | yes | the new window size for the folder action message to process |

### `edit action of`

Edit an action of a folder

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |
| `using action name` | `text` | yes | ...or the name of the action to edit |
| `using action number` | `integer` | yes | the index number of the action to edit... |

**Returns:** `file`

### `increment`

cause the target process to behave as if the UI element were incremented

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |

**Returns:** `UI element`

### `key down`

cause the target process to behave as if keys were held down

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text | integer | eMky | eMky` | no | a keystroke, key code, or (list of) modifier key names. |

### `key up`

cause the target process to behave as if keys were released

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text | integer | eMky | eMky` | no | a keystroke, key code, or (list of) modifier key names. |

### `pick`

cause the target process to behave as if the UI element were picked

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |

**Returns:** `UI element`

### `remove action from`

Remove a folder action from a folder

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |
| `using action name` | `text` | yes | ...or the name of the action to remove |
| `using action number` | `integer` | yes | the index number of the action to remove... |

**Returns:** `folder action`

### Enumerations

### `actn`

| Value | Description |
|-------|-------------|
| `items added` | items added |
| `items removed` | items removed |
| `window closed` | window closed |
| `window moved` | window moved |
| `window opened` | window opened |

---

## Scripting Definition Suite

Terms and Events for examining the System Events scripting definition

### Classes

### `scripting class`

A class within a suite within a scripting definition

*Plural:* `scripting classes`

**Contains:** `scripting element`, `scripting property`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the class |
| `id` | `text` | r/o | The unique identifier of the class |
| `description` | `text` | r/o | The description of the class |
| `hidden` | `boolean` | r/o | Is the class hidden? |
| `plural name` | `text` | r/o | The plural name of the class |
| `suite name` | `text` | r/o | The name of the suite to which this class belongs |
| `superclass` | `scripting class` | r/o | The class from which this class inherits |

### `scripting command`

A command within a suite within a scripting definition

*Plural:* `scripting commands`

**Contains:** `scripting parameter`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the command |
| `id` | `text` | r/o | The unique identifier of the command |
| `description` | `text` | r/o | The description of the command |
| `direct parameter` | `scripting parameter` | r/o | The direct parameter of the command |
| `hidden` | `boolean` | r/o | Is the command hidden? |
| `scripting result` | `scripting result object` | r/o | The object or data returned by this command |
| `suite name` | `text` | r/o | The name of the suite to which this command belongs |

### `scripting definition object`

The scripting definition of the System Events applicaation

**Contains:** `scripting suite`

### `scripting element`

An element within a class within a suite within a scripting definition

*Plural:* `scripting elements`

### `scripting enumeration`

An enumeration within a suite within a scripting definition

*Plural:* `scripting enumerations`

**Contains:** `scripting enumerator`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the enumeration |
| `id` | `text` | r/o | The unique identifier of the enumeration |
| `hidden` | `boolean` | r/o | Is the enumeration hidden? |

### `scripting enumerator`

An enumerator within an enumeration within a suite within a scripting definition

*Plural:* `scripting enumerators`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the enumerator |
| `id` | `text` | r/o | The unique identifier of the enumerator |
| `description` | `text` | r/o | The description of the enumerator |
| `hidden` | `boolean` | r/o | Is the enumerator hidden? |

### `scripting parameter`

A parameter within a command within a suite within a scripting definition

*Plural:* `scripting parameters`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the parameter |
| `id` | `text` | r/o | The unique identifier of the parameter |
| `description` | `text` | r/o | The description of the parameter |
| `hidden` | `boolean` | r/o | Is the parameter hidden? |
| `kind` | `text` | r/o | The kind of object or data specified by this parameter |
| `optional` | `boolean` | r/o | Is the parameter optional? |

### `scripting property`

A property within a class within a suite within a scripting definition

*Plural:* `scripting properties`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the property |
| `id` | `text` | r/o | The unique identifier of the property |
| `access` | `accs` | r/o | The type of access to this property |
| `description` | `text` | r/o | The description of the property |
| `enumerated` | `boolean` | r/o | Is the property's value an enumerator? |
| `hidden` | `boolean` | r/o | Is the property hidden? |
| `kind` | `text` | r/o | The kind of object or data returned by this property |
| `listed` | `boolean` | r/o | Is the property's value a list? |

### `scripting result object`

The result of a command within a suite within a scripting definition

*Plural:* `scripting result objects`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `description` | `text` | r/o | The description of the property |
| `enumerated` | `boolean` | r/o | Is the scripting result's value an enumerator? |
| `kind` | `text` | r/o | The kind of object or data returned by this property |
| `listed` | `boolean` | r/o | Is the scripting result's value a list? |

### `scripting suite`

A suite within a scripting definition

*Plural:* `scripting suites`

**Contains:** `scripting command`, `scripting class`, `scripting enumeration`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | The name of the suite |
| `id` | `text` | r/o | The unique identifier of the suite |
| `description` | `text` | r/o | The description of the suite |
| `hidden` | `boolean` | r/o | Is the suite hidden? |

### Enumerations

### `accs`

| Value | Description |
|-------|-------------|
| `none` | none |
| `read only` | read only |
| `read write` | read write |
| `write only` | write only |
