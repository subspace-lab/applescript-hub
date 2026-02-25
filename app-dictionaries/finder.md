# Finder AppleScript Dictionary

> Auto-generated from `Finder.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Finder"`

## Table of Contents

- [Standard Suite](#standard-suite)
- [Finder Basics](#finder-basics)
- [Finder items](#finder-items)
- [Containers and folders](#containers-and-folders)
- [Files](#files)
- [Window classes](#window-classes)
- [Legacy suite](#legacy-suite)
- [Type Definitions](#type-definitions)
- [Enumerations](#enumerations)

---

## Standard Suite

Common terms that most applications should support

### Commands

### `open`

Open the specified object(s)

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | list of objects to open |
| `using` | `specifier` | yes | the application file to open the object with |
| `with properties` | `record` | yes | the initial values for the properties, to be included with the open command sent to the application that opens the direct object |

### `print`

Print the specified object(s)

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | list of objects to print |
| `with properties` | `record` | yes | optional properties to be included with the print command sent to the application that prints the direct object |

### `quit`

Quit the Finder

### `activate`

Activate the specified window (or the Finder)

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | yes | the window to activate (if not specified, activates the Finder) |

### `close`

Close an object

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object to close |

### `count`

Return the number of elements of a particular class within an object

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object whose elements are to be counted |
| `each` | `type` | no | the class of the elements to be counted |

**Returns:** `integer` — the number of elements

### `data size`

Return the size in bytes of an object

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object whose data size is to be returned |
| `as` | `type` | yes | the data type for which the size is calculated |

**Returns:** `integer` — the size of the object in bytes

### `delete`

Move an item from its container to the trash

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the item to delete |

**Returns:** `specifier` — to the item that was just deleted

### `duplicate`

Duplicate one or more object(s)

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object(s) to duplicate |
| `to` | `location specifier` | yes | the new location for the object(s) |
| `replacing` | `boolean` | yes | Specifies whether or not to replace items in the destination that have the same name as items being duplicated |
| `routing suppressed` | `boolean` | yes | Specifies whether or not to autoroute items (default is false). Only applies when copying to the system folder. |
| `exact copy` | `boolean` | yes | Specifies whether or not to copy permissions/ownership as is |

**Returns:** `specifier` — to the duplicated object(s)

### `exists`

Verify if an object exists

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object in question |

**Returns:** `boolean` — true if it exists, false if not

### `make`

Make a new element

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `new` | `type` | no | the class of the new element |
| `at` | `location specifier` | no | the location at which to insert the element |
| `to` | `specifier` | yes | when creating an alias file, the original item to create an alias to or when creating a file viewer window, the target of the window |
| `with properties` | `record` | yes | the initial values for the properties of the element |

**Returns:** `specifier` — to the new object(s)

### `move`

Move object(s) to a new location

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object(s) to move |
| `to` | `location specifier` | no | the new location for the object(s) |
| `replacing` | `boolean` | yes | Specifies whether or not to replace items in the destination that have the same name as items being moved |
| `positioned at` | `list` | yes | Gives a list (in local window coordinates) of positions for the destination items |
| `routing suppressed` | `boolean` | yes | Specifies whether or not to autoroute items (default is false). Only applies when moving to the system folder. |

**Returns:** `specifier` — to the object(s) after they have been moved

### `select`

Select the specified object(s)

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object to select |

---

## Finder Basics

Commonly-used Finder commands and object classes

### Commands

### `openVirtualLocation`

Private event to open a virtual location

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | the location to open |

### `copy`

(NOT AVAILABLE YET) Copy the selected items to the clipboard (the Finder must be the front application)

### `sort`

Return the specified object(s) in a sorted list

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | a list of finder objects to sort |
| `by` | `property` | no | the property to sort the items by (name, index, date, etc.) |

**Returns:** `specifier` — the sorted items in their new order

### Classes

### `application`

The Finder

**Contains:** `item`, `container`, `disk`, `folder`, `file`, `alias file`, `application file`, `document file`, `internet location file`, `clipping`, `package`, `window`, `Finder window`, `clipping window`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `clipboard` | `specifier` | r/o | (NOT AVAILABLE YET) the Finder’s clipboard window |
| `name` | `text` | r/o | the Finder’s name |
| `visible` | `boolean` | r/w | Is the Finder’s layer visible? |
| `frontmost` | `boolean` | r/w | Is the Finder the frontmost process? |
| `selection` | `specifier` | r/w | the selection in the frontmost Finder window |
| `insertion location` | `specifier` | r/o | the container in which a new folder would appear if “New Folder” was selected |
| `product version` | `text` | r/o | the version of the System software running on this computer |
| `version` | `text` | r/o | the version of the Finder |
| `startup disk` | `disk` | r/o | the startup disk |
| `desktop` | `desktop-object` | r/o | the desktop |
| `trash` | `trash-object` | r/o | the trash |
| `home` | `folder` | r/o | the home directory |
| `computer container` | `computer-object` | r/o | the computer location (as in Go > Computer) |
| `Finder preferences` | `preferences` | r/o | Various preferences that apply to the Finder as a whole |

---

## Finder items

Commands used with file system items, and basic item definition

### Commands

### `clean up`

Arrange items in window nicely (only applies to open windows in icon view that are not kept arranged)

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the window to clean up |
| `by` | `property` | yes | the order in which to clean up the objects (name, index, date, etc.) |

### `eject`

Eject the specified disk(s)

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | yes | the disk(s) to eject |

### `empty`

Empty the trash

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | yes | “empty” and “empty trash” both do the same thing |
| `security` | `boolean` | yes | (obsolete) |

### `erase`

(NOT AVAILABLE) Erase the specified disk(s)

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the items to erase |

### `reveal`

Bring the specified object(s) into view

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object to be made visible |

### `update`

Update the display of the specified object(s) to match their on-disk representation

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the item to update |
| `necessity` | `boolean` | yes | only update if necessary (i.e. a finder window is open). default is false |
| `registering applications` | `boolean` | yes | register applications. default is true |

### Classes

### `item`

An item

*Plural:* `items`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/w | the name of the item |
| `displayed name` | `text` | r/o | the user-visible name of the item |
| `name extension` | `text` | r/w | the name extension of the item (such as “txt”) |
| `extension hidden` | `boolean` | r/w | Is the item's extension hidden from the user? |
| `index` | `integer` | r/o | the index in the front-to-back ordering within its container |
| `container` | `specifier` | r/o | the container of the item |
| `disk` | `specifier` | r/o | the disk on which the item is stored |
| `position` | `point` | r/w | the position of the item within its parent window (can only be set for an item in a window viewed as icons or buttons) |
| `desktop position` | `point` | r/w | the position of the item on the desktop |
| `bounds` | `rectangle` | r/w | the bounding rectangle of the item (can only be set for an item in a window viewed as icons or buttons) |
| `label index` | `integer` | r/w | the label of the item |
| `locked` | `boolean` | r/w | Is the file locked? |
| `kind` | `text` | r/o | the kind of the item |
| `description` | `text` | r/o | a description of the item |
| `comment` | `text` | r/w | the comment of the item, displayed in the “Get Info” window |
| `size` | `double integer` | r/o | the logical size of the item |
| `physical size` | `double integer` | r/o | the actual space used by the item on disk |
| `creation date` | `date` | r/o | the date on which the item was created |
| `modification date` | `date` | r/w | the date on which the item was last modified |
| `icon` | `icon family` | r/w | the icon bitmap of the item |
| `URL` | `text` | r/o | the URL of the item |
| `owner` | `text` | r/w | the user that owns the container |
| `group` | `text` | r/w | the user or group that has special access to the container |
| `owner privileges` | `priv` | r/w |  |
| `group privileges` | `priv` | r/w |  |
| `everyones privileges` | `priv` | r/w |  |
| `information window` | `specifier` | r/o | the information window for the item |
| `properties` | `record` | r/w | every property of an item |
| `class` | `type` | r/o | the class of the item |

### Enumerations

### `priv`

| Value | Description |
|-------|-------------|
| `read only` |  |
| `read write` |  |
| `write only` |  |
| `none` |  |

---

## Containers and folders

Classes that can contain other file system items

### Classes

### `container`

An item that contains other items

*Plural:* `containers`

**Contains:** `item`, `container`, `folder`, `file`, `alias file`, `application file`, `document file`, `internet location file`, `clipping`, `package`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `entire contents` | `specifier` | r/o | the entire contents of the container, including the contents of its children |
| `expandable` | `boolean` | r/o | (NOT AVAILABLE YET) Is the container capable of being expanded as an outline? |
| `expanded` | `boolean` | r/w | (NOT AVAILABLE YET) Is the container opened as an outline? (can only be set for containers viewed as lists) |
| `completely expanded` | `boolean` | r/w | (NOT AVAILABLE YET) Are the container and all of its children opened as outlines? (can only be set for containers viewed as lists) |
| `container window` | `specifier` | r/o | the container window for this folder |

### `computer-object`

the Computer location (as in Go > Computer)

### `disk`

A disk

*Plural:* `disks`

**Contains:** `item`, `container`, `folder`, `file`, `alias file`, `application file`, `document file`, `internet location file`, `clipping`, `package`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `integer` | r/o | the unique id for this disk (unchanged while disk remains connected and Finder remains running) |
| `capacity` | `double integer` | r/o | the total number of bytes (free or used) on the disk |
| `free space` | `double integer` | r/o | the number of free bytes left on the disk |
| `ejectable` | `boolean` | r/o | Can the media be ejected (floppies, CDs, and so on)? |
| `local volume` | `boolean` | r/o | Is the media a local volume (as opposed to a file server)? |
| `startup` | `boolean` | r/o | Is this disk the boot disk? |
| `format` | `edfm` | r/o | the filesystem format of this disk |
| `journaling enabled` | `boolean` | r/o | Does this disk do file system journaling? |
| `ignore privileges` | `boolean` | r/w | Ignore permissions on this disk? |

### `folder`

A folder

*Plural:* `folders`

**Contains:** `item`, `container`, `folder`, `file`, `alias file`, `application file`, `document file`, `internet location file`, `clipping`, `package`

### `desktop-object`

Desktop-object is the class of the “desktop” object

**Contains:** `item`, `container`, `disk`, `folder`, `file`, `alias file`, `application file`, `document file`, `internet location file`, `clipping`, `package`

### `trash-object`

Trash-object is the class of the “trash” object

**Contains:** `item`, `container`, `folder`, `file`, `alias file`, `application file`, `document file`, `internet location file`, `clipping`, `package`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `warns before emptying` | `boolean` | r/w | Display a dialog when emptying the trash? |

### Enumerations

### `edfm`

| Value | Description |
|-------|-------------|
| `Mac OS format` |  |
| `Mac OS Extended format` |  |
| `UFS format` |  |
| `NFS format` |  |
| `audio format` |  |
| `ProDOS format` |  |
| `MSDOS format` |  |
| `NTFS format` |  |
| `ISO 9660 format` |  |
| `High Sierra format` |  |
| `QuickTake format` |  |
| `Apple Photo format` |  |
| `AppleShare format` |  |
| `UDF format` |  |
| `WebDAV format` |  |
| `FTP format` |  |
| `Packet written UDF format` |  |
| `Xsan format` |  |
| `APFS format` |  |
| `ExFAT format` |  |
| `SMB format` |  |
| `unknown format` |  |

---

## Files

Classes representing files

### Classes

### `file`

A file

*Plural:* `files`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `file type` | `type` | r/w | the OSType identifying the type of data contained in the item |
| `creator type` | `type` | r/w | the OSType identifying the application that created the item |
| `stationery` | `boolean` | r/w | Is the file a stationery pad? |
| `product version` | `text` | r/o | the version of the product (visible at the top of the “Get Info” window) |
| `version` | `text` | r/o | the version of the file (visible at the bottom of the “Get Info” window) |

### `alias file`

An alias file (created with “Make Alias”)

*Plural:* `alias files`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `original item` | `specifier` | r/w | the original item pointed to by the alias |

### `application file`

An application's file on disk

*Plural:* `application files`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | the bundle identifier or creator type of the application |
| `suggested size` | `integer` | r/o | (AVAILABLE IN 10.1 TO 10.4) the memory size with which the developer recommends the application be launched |
| `minimum size` | `integer` | r/w | (AVAILABLE IN 10.1 TO 10.4) the smallest memory size with which the application can be launched |
| `preferred size` | `integer` | r/w | (AVAILABLE IN 10.1 TO 10.4) the memory size with which the application will be launched |
| `accepts high level events` | `boolean` | r/o | Is the application high-level event aware? (OBSOLETE: always returns true) |
| `has scripting terminology` | `boolean` | r/o | Does the process have a scripting terminology, i.e., can it be scripted? |
| `opens in Classic` | `boolean` | r/w | (AVAILABLE IN 10.1 TO 10.4) Should the application launch in the Classic environment? |

### `document file`

A document file

*Plural:* `document files`

### `internet location file`

A file containing an internet location

*Plural:* `internet location files`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `location` | `text` | r/o | the internet location |

### `clipping`

A clipping

*Plural:* `clippings`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `clipping window` | `specifier` | r/o | (NOT AVAILABLE YET) the clipping window for this clipping |

### `package`

A package

*Plural:* `packages`

---

## Window classes

Classes representing windows

### Classes

### `window`

A window

*Plural:* `windows`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `integer` | r/o | the unique id for this window |
| `position` | `point` | r/w | the upper left position of the window |
| `bounds` | `rectangle` | r/w | the boundary rectangle for the window |
| `titled` | `boolean` | r/o | Does the window have a title bar? |
| `name` | `text` | r/o | the name of the window |
| `index` | `integer` | r/w | the number of the window in the front-to-back layer ordering |
| `closeable` | `boolean` | r/o | Does the window have a close box? |
| `floating` | `boolean` | r/o | Does the window have a title bar? |
| `modal` | `boolean` | r/o | Is the window modal? |
| `resizable` | `boolean` | r/o | Is the window resizable? |
| `zoomable` | `boolean` | r/o | Is the window zoomable? |
| `zoomed` | `boolean` | r/w | Is the window zoomed? |
| `visible` | `boolean` | r/o | Is the window visible (always true for open Finder windows)? |
| `collapsed` | `boolean` | r/w | Is the window collapsed |
| `properties` | `record` | r/w | every property of a window |

### `Finder window`

A file viewer window

*Plural:* `Finder windows`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `target` | `specifier` | r/w | the container at which this file viewer is targeted |
| `current view` | `ecvw` | r/w | the current view for the container window |
| `icon view options` | `icon view options` | r/o | the icon view options for the container window |
| `list view options` | `list view options` | r/o | the list view options for the container window |
| `column view options` | `column view options` | r/o | the column view options for the container window |
| `toolbar visible` | `boolean` | r/w | Is the window's toolbar visible? |
| `statusbar visible` | `boolean` | r/w | Is the window's status bar visible? |
| `pathbar visible` | `boolean` | r/w | Is the window's path bar visible? |
| `sidebar width` | `integer` | r/w | the width of the sidebar for the container window |

### `desktop window`

the desktop window

### `information window`

An inspector window (opened by “Show Info”)

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `item` | `specifier` | r/o | the item from which this window was opened |
| `current panel` | `ipnl` | r/w | the current panel in the information window |

### `preferences window`

The Finder Preferences window

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `current panel` | `pple` | r/w | The current panel in the Finder preferences window |

### `clipping window`

The window containing a clipping

*Plural:* `clipping windows`

### Enumerations

### `ipnl`

| Value | Description |
|-------|-------------|
| `General Information panel` |  |
| `Sharing panel` |  |
| `Memory panel` |  |
| `Preview panel` |  |
| `Application panel` |  |
| `Languages panel` |  |
| `Plugins panel` |  |
| `Name & Extension panel` |  |
| `Comments panel` |  |
| `Content Index panel` |  |
| `Burning panel` |  |
| `More Info panel` |  |
| `Simple Header panel` |  |

### `pple`

| Value | Description |
|-------|-------------|
| `General Preferences panel` |  |
| `Label Preferences panel` |  |
| `Sidebar Preferences panel` |  |
| `Advanced Preferences panel` |  |

### `ecvw`

| Value | Description |
|-------|-------------|
| `icon view` |  |
| `list view` |  |
| `column view` |  |
| `group view` |  |
| `flow view` |  |

---

## Legacy suite

Operations formerly handled by the Finder, but now automatically delegated to other applications

### Commands

### `restart`

Restart the computer

### `shut down`

Shut Down the computer

### `sleep`

Put the computer to sleep

### Classes

### `process`

A process running on this computer

*Plural:* `processes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/o | the name of the process |
| `visible` | `boolean` | r/w | Is the process' layer visible? |
| `frontmost` | `boolean` | r/w | Is the process the frontmost process? |
| `file` | `specifier` | r/o | the file from which the process was launched |
| `file type` | `type` | r/o | the OSType of the file type of the process |
| `creator type` | `type` | r/o | the OSType of the creator of the process (the signature) |
| `accepts high level events` | `boolean` | r/o | Is the process high-level event aware (accepts open application, open document, print document, and quit)? |
| `accepts remote events` | `boolean` | r/o | Does the process accept remote events? |
| `has scripting terminology` | `boolean` | r/o | Does the process have a scripting terminology, i.e., can it be scripted? |
| `total partition size` | `integer` | r/o | the size of the partition with which the process was launched |
| `partition space used` | `integer` | r/o | the number of bytes currently used in the process' partition |

### `application process`

A process launched from an application file

*Plural:* `application processes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `application file` | `application file` | r/o | the application file from which this process was launched |

### `desk accessory process`

A process launched from a desk accessory file

*Plural:* `desk accessory processes`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `desk accessory file` | `specifier` | r/o | the desk accessory file from which this process was launched |

---

## Type Definitions

Definitions of records used in scripting the Finder

### Classes

### `preferences`

The Finder Preferences

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `window` | `preferences window` | r/o | the window that would open if Finder preferences was opened |
| `icon view options` | `icon view options` | r/o | the default icon view options |
| `list view options` | `list view options` | r/o | the default list view options |
| `column view options` | `column view options` | r/o | the column view options for all windows |
| `folders spring open` | `boolean` | r/w | Spring open folders after the specified delay? |
| `delay before springing` | `real` | r/w | the delay before springing open a container in seconds (from 0.167 to 1.169) |
| `desktop shows hard disks` | `boolean` | r/w | Hard disks appear on the desktop? |
| `desktop shows external hard disks` | `boolean` | r/w | External hard disks appear on the desktop? |
| `desktop shows removable media` | `boolean` | r/w | CDs, DVDs, and iPods appear on the desktop? |
| `desktop shows connected servers` | `boolean` | r/w | Connected servers appear on the desktop? |
| `new window target` | `specifier` | r/w | target location for a newly-opened Finder window |
| `folders open in new windows` | `boolean` | r/w | Folders open into new windows? |
| `folders open in new tabs` | `boolean` | r/w | Folders open into new tabs? |
| `new windows open in column view` | `boolean` | r/w | Open new windows in column view? |
| `all name extensions showing` | `boolean` | r/w | Show name extensions, even for items whose “extension hidden” is true? |

### `label`

(NOT AVAILABLE YET) A Finder label (name and color)

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/w | the name associated with the label |
| `index` | `integer` | r/w | the index in the front-to-back ordering within its container |
| `color` | `RGB color` | r/w | the color associated with the label |

### `icon family`

(NOT AVAILABLE YET) A family of icons

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `large monochrome icon and mask` | `ICN#` | r/o | the large black-and-white icon and the mask for large icons |
| `large 8 bit mask` | `l8mk` | r/o | the large 8-bit mask for large 32-bit icons |
| `large 32 bit icon` | `il32` | r/o | the large 32-bit color icon |
| `large 8 bit icon` | `icl8` | r/o | the large 8-bit color icon |
| `large 4 bit icon` | `icl4` | r/o | the large 4-bit color icon |
| `small monochrome icon and mask` | `ics#` | r/o | the small black-and-white icon and the mask for small icons |
| `small 8 bit mask` | `s8mk` | r/o | the small 8-bit mask for small 32-bit icons |
| `small 32 bit icon` | `is32` | r/o | the small 32-bit color icon |
| `small 8 bit icon` | `ics8` | r/o | the small 8-bit color icon |
| `small 4 bit icon` | `ics4` | r/o | the small 4-bit color icon |

### `icon view options`

the icon view options

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `arrangement` | `earr` | r/w | the property by which to keep icons arranged |
| `icon size` | `integer` | r/w | the size of icons displayed in the icon view |
| `shows item info` | `boolean` | r/w | additional info about an item displayed in icon view |
| `shows icon preview` | `boolean` | r/w | displays a preview of the item in icon view |
| `text size` | `integer` | r/w | the size of the text displayed in the icon view |
| `label position` | `epos` | r/w | the location of the label in reference to the icon |
| `background picture` | `file` | r/w | the background picture of the icon view |
| `background color` | `RGB color` | r/w | the background color of the icon view |

### `column view options`

the column view options

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `text size` | `integer` | r/w | the size of the text displayed in the column view |
| `shows icon` | `boolean` | r/w | displays an icon next to the label in column view |
| `shows icon preview` | `boolean` | r/w | displays a preview of the item in column view |
| `shows preview column` | `boolean` | r/w | displays the preview column in column view |
| `discloses preview pane` | `boolean` | r/w | discloses the preview pane of the preview column in column view |

### `list view options`

the list view options

**Contains:** `column`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `calculates folder sizes` | `boolean` | r/w | Are folder sizes calculated and displayed in the window? |
| `shows icon preview` | `boolean` | r/w | displays a preview of the item in list view |
| `icon size` | `lvic` | r/w | the size of icons displayed in the list view |
| `text size` | `integer` | r/w | the size of the text displayed in the list view |
| `sort column` | `column` | r/w | the column that the list view is sorted on |
| `uses relative dates` | `boolean` | r/w | Are relative dates (e.g., today, yesterday) shown in the list view? |

### `column`

a column of a list view

*Plural:* `columns`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `index` | `integer` | r/w | the index in the front-to-back ordering within its container |
| `name` | `elsv` | r/o | the column name |
| `sort direction` | `sodr` | r/w | The direction in which the window is sorted |
| `width` | `integer` | r/w | the width of this column |
| `minimum width` | `integer` | r/o | the minimum allowed width of this column |
| `maximum width` | `integer` | r/o | the maximum allowed width of this column |
| `visible` | `boolean` | r/w | is this column visible |

### `alias list`

A list of aliases. Use ‘as alias list’ when a list of aliases is needed (instead of a list of file system item references).

### Enumerations

### `earr`

| Value | Description |
|-------|-------------|
| `not arranged` |  |
| `snap to grid` |  |
| `arranged by name` |  |
| `arranged by modification date` |  |
| `arranged by creation date` |  |
| `arranged by size` |  |
| `arranged by kind` |  |
| `arranged by label` |  |

### `epos`

| Value | Description |
|-------|-------------|
| `right` |  |
| `bottom` |  |

### `sodr`

| Value | Description |
|-------|-------------|
| `normal` |  |
| `reversed` |  |

### `elsv`

| Value | Description |
|-------|-------------|
| `name column` |  |
| `modification date column` |  |
| `creation date column` |  |
| `size column` |  |
| `kind column` |  |
| `label column` |  |
| `version column` |  |
| `comment column` |  |

### `lvic`

| Value | Description |
|-------|-------------|
| `small icon` |  |
| `large icon` |  |

---

## Enumerations

Enumerations for the Finder

### Enumerations

### `isiz`

| Value | Description |
|-------|-------------|
| `mini` |  |
| `small` |  |
| `large` |  |

### `sort`

| Value | Description |
|-------|-------------|
| `name` |  |
| `modification date` |  |
| `creation date` |  |
| `size` |  |
| `kind` |  |
| `label index` |  |
| `comment` |  |
| `version` |  |
