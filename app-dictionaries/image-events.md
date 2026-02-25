# Imageevents AppleScript Dictionary

> Auto-generated from `ImageEvents.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Imageevents"`

## Table of Contents

- [Disk-Folder-File Suite](#disk-folder-file-suite)
- [Image Suite](#image-suite)
- [Image Events Suite](#image-events-suite)

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

**Returns:** `any | alias | alias`

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

**Responds to:** `open`

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

## Image Suite

Terms and Events for controlling Images

### Commands

### `close`

Close an image

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |
| `saving` | `savo` | yes | Specifies whether changes should be saved before closing. |
| `saving in` | `text` | yes | The file in which to save the image. |

### `crop`

Crop an image

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |
| `to dimensions` | `integer` | no | the width and height of the new image, respectively, in pixels, as a pair of integers |

### `embed`

Embed an image with an ICC profile

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |
| `with source` | `profile` | no | the profile to embed in the image |

### `flip`

Flip an image

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |
| `horizontal` | `boolean` | yes | flip horizontally |
| `vertical` | `boolean` | yes | flip vertically |

### `match`

Match an image

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |
| `to destination` | `profile` | no | the destination profile for the match |

### `pad`

Pad an image

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |
| `to dimensions` | `integer` | no | the width and height of the new image, respectively, in pixels, as a pair of integers |
| `with pad color` | `integer` | yes | the RGB color values with which to pad the new image, as a list of integers |

### `rotate`

Rotate an image

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |
| `to angle` | `real` | no | rotate using an angle |

### `save`

Save an image to a file in one of various formats

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |
| `as` | `typv` | yes | file type in which to save the image ( default is to make no change ) |
| `icon` | `boolean` | yes | Shall an icon be added? ( default is false ) |
| `in` | `text` | yes | file path in which to save the image, in HFS or POSIX form |
| `PackBits` | `boolean` | yes | Are the bytes to be compressed with PackBits? ( default is false, applies only to TIFF ) |
| `with compression level` | `cmlv` | yes | specifies the compression level of the resultant file ( applies only to JPEG ) |

**Returns:** `alias`

### `scale`

Scale an image

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |
| `by factor` | `real` | yes | scale using a scalefactor |
| `to size` | `integer` | yes | scale using a max width/length |

### `unembed`

Remove any embedded ICC profiles from an image

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the object for the command |

### Classes

### `display`

A monitor connected to the computer

*Plural:* `displays`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `display number` | `integer` | r/o | the number of the display |
| `display profile` | `profile` | r/o | the profile for the display |
| `name` | `text` | r/o | the name of the display |

### `image`

An image contained in a file

*Plural:* `images`

**Contains:** `metadata tag`, `profile`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `bit depth` | `bitz` | r/o | bit depth of the image's color representation |
| `color space` | `pSpc` | r/o | color space of the image's color representation |
| `dimensions` | `integer` | r/o | the width and height of the image, respectively, in pixels |
| `embedded profile` | `profile` | r/o | the profile, if any, embedded in the image |
| `file type` | `typz | any` | r/o | file type of the image's file |
| `image file` | `file` | r/o | the file that contains the image |
| `location` | `disk item` | r/o | the folder or disk that encloses the file that contains the image |
| `name` | `text` | r/o | the name of the image |
| `resolution` | `real` | r/o | the horizontal and vertical pixel density of the image, respectively, in dots per inch |

**Responds to:** `close`, `crop`, `embed`, `flip`, `match`, `pad`, `rotate`, `save`, `scale`, `unembed`

### `metadata tag`

A metadata tag: EXIF, IPTC, etc.

*Plural:* `metadata tags`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `description` | `text` | r/o | the description of the tag's function |
| `name` | `text` | r/o | the name of the tag |
| `value` | `boolean | integer | real | text | profile | any` | r/o | the current setting of the tag |

### `profile`

A ColorSync ICC profile.

*Plural:* `profiles`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `color space` | `pSpc` | r/o | the color space of the profile |
| `connection space` | `pPCS` | r/o | the connection space of the profile |
| `creation date` | `date` | r/o | the creation date of the profile |
| `creator` | `text` | r/o | the creator type of the profile |
| `device class` | `pCla` | r/o | the device class of the profile |
| `device manufacturer` | `text` | r/o | the device manufacturer of the profile |
| `device model` | `integer` | r/o | the device model of the profile |
| `location` | `alias | any` | r/o | the file location of the profile |
| `name` | `text` | r/o | the description text of the profile |
| `platform` | `text` | r/o | the intended platform of the profile |
| `preferred CMM` | `text` | r/o | the preferred CMM of the profile |
| `quality` | `pQua` | r/o | the quality of the profile |
| `rendering intent` | `pRdr` | r/o | the rendering intent of the profile |
| `size` | `integer` | r/o | the size of the profile in bytes |
| `version` | `text` | r/o | the version number of the profile |

### Enumerations

### `bitz`

| Value | Description |
|-------|-------------|
| `best` | best |
| `black & white` | black & white |
| `color` | color |
| `four colors` | four colors |
| `four grays` | four grays |
| `grayscale` | grayscale |
| `millions of colors` | millions of colors |
| `millions of colors plus` | millions of colors plus |
| `sixteen colors` | sixteen colors |
| `sixteen grays` | sixteen grays |
| `thousands of colors` | thousands of colors |
| `two hundred fifty six colors` | two hundred fifty six colors |
| `two hundred fifty six grays` | two hundred fifty six grays |

### `pCla`

| Value | Description |
|-------|-------------|
| `abstract` | abstract profile |
| `colorspace` | colorspace profile |
| `input` | input device |
| `link` | device-link profile |
| `monitor` | display device |
| `named` | named color space profile |
| `output` | output device |

### `pPCS`

| Value | Description |
|-------|-------------|
| `Lab` | Lab |
| `XYZ` | XYZ |

### `cmlv`

| Value | Description |
|-------|-------------|
| `high` | High compression |
| `low` | Low compression |
| `medium` | Medium compression |

### `typz`

| Value | Description |
|-------|-------------|
| `BMP` | BMP |
| `GIF` | GIF |
| `JPEG` | JPEG |
| `JPEG2` | JPEG2 |
| `MacPaint` | MacPaint |
| `PDF` | PDF |
| `Photoshop` | Photoshop |
| `PICT` | PICT |
| `PNG` | PNG |
| `PSD` | PSD |
| `QuickTime Image` | QuickTime Image |
| `SGI` | SGI |
| `Text` | Text |
| `TGA` | TGA |
| `TIFF` | TIFF |

### `pQua`

| Value | Description |
|-------|-------------|
| `best` | best |
| `draft` | draft |
| `normal` | normal |

### `pSpc`

| Value | Description |
|-------|-------------|
| `CMYK` | CMYK |
| `Eight channel` | Eight channel |
| `Eight color` | Eight color |
| `Five channel` | Five channel |
| `Five color` | Five color |
| `Gray` | Gray |
| `Lab` | Lab |
| `Named` | Named |
| `RGB` | RGB |
| `Seven channel` | Seven channel |
| `Seven color` | Seven color |
| `Six channel` | Six channel |
| `Six color` | Six color |
| `XYZ` | XYZ |

### `pRdr`

| Value | Description |
|-------|-------------|
| `absolute colorimetric intent` | absolute colorimetric |
| `perceptual intent` | perceptual |
| `relative colorimetric intent` | relative colorimetric |
| `saturation intent` | saturation |

### `savo`

| Value | Description |
|-------|-------------|
| `no` | Do not save the image. |
| `yes` | Save the image. |

### `qual`

| Value | Description |
|-------|-------------|
| `best` | best |
| `high` | high |
| `least` | least |
| `low` | low |
| `medium` | medium |

### `typv`

| Value | Description |
|-------|-------------|
| `BMP` | BMP |
| `JPEG` | JPEG |
| `JPEG2` | JPEG2 |
| `PICT` | PICT |
| `PNG` | PNG |
| `PSD` | PSD |
| `QuickTime Image` | QuickTime Image |
| `TIFF` | TIFF |

---

## Image Events Suite

Terms and Events for controlling the Image Events application

### Enumerations

### `saveable file format`

| Value | Description |
|-------|-------------|
| `text` | Text File Format |
