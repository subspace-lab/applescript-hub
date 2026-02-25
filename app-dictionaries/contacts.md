# Contacts AppleScript Dictionary

> Auto-generated from `Contacts.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Contacts"`

## Table of Contents

- [Standard Suite](#standard-suite)
- [Contacts Script Suite](#contacts-script-suite)
- [Address Book Rollover Suite](#address-book-rollover-suite)

---

## Standard Suite

Common classes and commands for all applications.

### Commands

### `make`

Create a new object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `new` | `type` | no | The class of the new object. |
| `at` | `location specifier` | yes | The location at which to insert the object. |
| `with data` | `any` | yes | The initial contents of the object. |
| `with properties` | `record` | yes | The initial values for properties of the object. |

**Returns:** `specifier` — The new object.

---

## Contacts Script Suite

commands and classes for Contacts scripting.

### Commands

### `add`

Add a child object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `entry` | no | object to add. |
| `to` | `specifier` | no | where to add this child to. |

**Returns:** `person`

### `remove`

Remove a child object.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `entry` | no | object to remove. |
| `from` | `specifier` | no | where to remove this child from. |

**Returns:** `person`

### `save`

Save all Contacts changes. Also see the unsaved property for the application class.

**Returns:** `any`

### Classes

### `address`

Address for the given record.

*Plural:* `addresses`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `city` | `text | missing value` | r/w | City part of the address. |
| `formatted address` | `text | missing value` | r/o | properly formatted string for this address. |
| `street` | `text | missing value` | r/w | Street part of the address, multiple lines separated by carriage returns. |
| `id` | `text` | r/w | unique identifier for this address. |
| `zip` | `text | missing value` | r/w | Zip or postal code of the address. |
| `country` | `text | missing value` | r/w | Country part of the address. |
| `label` | `text | missing value` | r/w | Label. |
| `country code` | `text | missing value` | r/w | Country code part of the address (should be a two character iso country code). |
| `state` | `text | missing value` | r/w | State, Province, or Region part of the address. |

### `AIM Handle`

User name for America Online (AOL) instant messaging.

*Plural:* `AIM handles`

### `contact info`

Container object in the database, holds a key and a value

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `label` | `text | missing value` | r/w | Label is the label associated with value like "work", "home", etc. |
| `value` | `text | date | missing value` | r/w | Value. |
| `id` | `text` | r/o | unique identifier for this entry, this is persistent, and stays with the record. |

### `custom date`

Arbitrary date associated with this person.

*Plural:* `custom dates`

### `email`

Email address for a person.

*Plural:* `emails`

### `entry`

An entry in the address book database

*Plural:* `entries`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `modification date` | `date` | r/o | when the contact was last modified. |
| `creation date` | `date` | r/o | when the contact was created. |
| `id` | `text` | r/o | unique and persistent identifier for this record. |
| `selected` | `boolean` | r/w | Is the entry selected? |

### `group`

A Group Record in the address book database

*Plural:* `groups`

**Contains:** `group`, `person`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/w | The name of this group. |

### `ICQ handle`

User name for ICQ instant messaging.

*Plural:* `ICQ handles`

### `instant message`

Address for instant messaging.

*Plural:* `instant messages`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `service name` | `text | missing value` | r/o | The service name of this instant message address. |
| `service type` | `instant message service type | missing value` | r/w | The service type of this instant message address. |
| `user name` | `text | missing value` | r/w | The user name of this instant message address. |

### `Jabber handle`

User name for Jabber instant messaging.

*Plural:* `Jabber handles`

### `MSN handle`

User name for Microsoft Network (MSN) instant messaging.

*Plural:* `MSN handles`

### `person`

A person in the address book database.

*Plural:* `people`

**Contains:** `MSN handle`, `url`, `address`, `phone`, `Jabber handle`, `group`, `custom date`, `AIM Handle`, `Yahoo handle`, `ICQ handle`, `instant message`, `social profile`, `related name`, `email`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `nickname` | `text | missing value` | r/w | The Nickname of this person. |
| `organization` | `text | missing value` | r/w | Organization that employs this person. |
| `maiden name` | `text | missing value` | r/w | The Maiden name of this person. |
| `suffix` | `text | missing value` | r/w | The Suffix of this person. |
| `vcard` | `text | missing value` | r/o | Person information in vCard format, this always returns a card in version 3.0 format. |
| `home page` | `text | missing value` | r/w | The home page of this person. |
| `birth date` | `date | missing value` | r/w | The birth date of this person. |
| `phonetic last name` | `text | missing value` | r/w | The phonetic version of the Last name of this person. |
| `title` | `text | missing value` | r/w | The title of this person. |
| `phonetic middle name` | `text | missing value` | r/w | The Phonetic version of the Middle name of this person. |
| `department` | `text | missing value` | r/w | Department that this person works for. |
| `image` | `TIFF picture | missing value` | r/w | Image for person. |
| `name` | `text` | r/o | First/Last name of the person, uses the name display order preference setting in Contacts. |
| `note` | `text | missing value` | r/w | Notes for this person. |
| `company` | `boolean` | r/w | Is the current record a company or a person. |
| `middle name` | `text | missing value` | r/w | The Middle name of this person. |
| `phonetic first name` | `text | missing value` | r/w | The phonetic version of the First name of this person. |
| `job title` | `text | missing value` | r/w | The job title of this person. |
| `last name` | `text | missing value` | r/w | The Last name of this person. |
| `first name` | `text | missing value` | r/w | The First name of this person. |

**Responds to:** `add`, `remove`

### `phone`

Phone number for a person.

*Plural:* `phones`

### `related name`

Other names related to this person.

*Plural:* `related names`

### `social profile`

Profile for social networks.

*Plural:* `social profiles`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `id` | `text` | r/o | The persistent unique identifier for this profile. |
| `service name` | `text | missing value` | r/w | The service name of this social profile. |
| `user name` | `text | missing value` | r/w | The username used with this social profile. |
| `user identifier` | `text | missing value` | r/w | A service-specific identifier used with this social profile. |
| `url` | `text | missing value` | r/w | The URL of this social profile. |

### `url`

URLs for this person.

*Plural:* `urls`

### `Yahoo handle`

User name for Yahoo instant messaging.

*Plural:* `Yahoo handles`

### Enumerations

### `saveable file format`

| Value | Description |
|-------|-------------|
| `archive` | The native Contacts file format |

### `instant message service type`

| Value | Description |
|-------|-------------|
| `AIM` |  |
| `Facebook` |  |
| `Gadu Gadu` |  |
| `Google Talk` |  |
| `ICQ` |  |
| `Jabber` |  |
| `MSN` |  |
| `QQ` |  |
| `Skype` |  |
| `Yahoo` |  |

---

## Address Book Rollover Suite

These event definitions are used for constructing Address Book Rollover plug-ins. They would not normally appear in a typical end user script.

### Commands

### `action property`

RollOver - Which property this roll over is associated with (Properties can be one of maiden name, phone, email, url, birth date, custom date, related name, aim, icq, jabber, msn, yahoo, address.)

**Returns:** `text`

### `action title`

RollOver - Returns the title that will be placed in the menu for this roll over

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `with` | `any` | no | property that that was returned from the "action property" handler. |
| `for` | `person` | no | Currently selected person. |

**Returns:** `text`

### `perform action`

RollOver - Performs the action on the given person and value

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `with` | `any` | no | property that that was returned from the "action property" handler. |
| `for` | `person` | no | Currently selected person. |

**Returns:** `boolean`

### `should enable action`

RollOver - Determines if the rollover action should be enabled for the given person and value

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `with` | `any` | no | property that that was returned from the "action property" handler. |
| `for` | `person` | no | Currently selected person. |

**Returns:** `boolean`
