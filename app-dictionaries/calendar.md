# Ical AppleScript Dictionary

> Auto-generated from `iCal.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Ical"`

## Table of Contents

- [Standard Suite](#standard-suite)
- [iCal](#ical)

---

## Standard Suite

Common classes and commands for all applications.

---

## iCal

iCal classes and commands

### Commands

### `create calendar`

Creates a new calendar (obsolete, will be removed in next release)

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `with name` | `text` | yes | the calendar new name |

### `reload calendars`

Tell the application to reload all calendar files contents

### `switch view`

Show calendar on the given view

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `to` | `view type` | no | the calendar view to be displayed |

### `view calendar`

Show calendar on the given date

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `at` | `date` | no | the date to be displayed |

### `GetURL`

Subscribe to a remote calendar through a webcal or http URL

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `text` | no | the iCal URL |

### `show`

Show the event or to-do in the calendar window

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `(direct)` | `specifier` | no | the item |

### `make`

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `new` | `type` | no | The class of the new object. |
| `at` | `location specifier` | yes | The location at which to insert the object. |
| `with data` | `any` | yes | The initial contents of the object. |
| `with properties` | `record` | yes | The initial values for properties of the object. |

**Returns:** `specifier` — The new object.

### `save`

### Classes

### `calendar`

This class represents a calendar.

**Contains:** `event`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `name` | `text` | r/w | This is the calendar title. |
| `title` | `text` | r/w | This is the calendar title. |
| `color` | `RGB color` | r/w | The calendar color. |
| `calendarIdentifier` | `text` | r/o | An unique calendar key |
| `writable` | `boolean` | r/o | This is the calendar title. |
| `description` | `text` | r/w | This is the calendar description. |

### `display alarm`

This class represents a message alarm.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `trigger interval` | `integer` | r/w | The interval in minutes between the event and the alarm: (positive for alarm that trigger after the event date or negative for alarms that trigger before). |
| `trigger date` | `date` | r/w | An absolute alarm date. |

### `mail alarm`

This class represents a mail alarm.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `trigger interval` | `integer` | r/w | The interval in minutes between the event and the alarm: (positive for alarm that trigger after the event date or negative for alarms that trigger before). |
| `trigger date` | `date` | r/w | An absolute alarm date. |

### `sound alarm`

This class represents a sound alarm.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `trigger interval` | `integer` | r/w | The interval in minutes between the event and the alarm: (positive for alarm that trigger after the event date or negative for alarms that trigger before). |
| `trigger date` | `date` | r/w | An absolute alarm date. |
| `sound name` | `text` | r/w | The system sound name to be used for the alarm |
| `sound file` | `text` | r/w | The (POSIX) path to the sound file to be used for the alarm |

### `open file alarm`

This class represents an 'open file' alarm. Starting with OS X 10.14, it is not possible to create new open file alarms or view URLs for existing open file alarms. Trying to save or modify an open file alarm will result in a save error. Editing other aspects of events or reminders that have existing open file alarms is allowed as long as the alarm isn't modified.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `trigger interval` | `integer` | r/w | The interval in minutes between the event and the alarm: (positive for alarm that trigger after the event date or negative for alarms that trigger before). |
| `trigger date` | `date` | r/w | An absolute alarm date. |
| `filepath` | `text` | r/w | The (POSIX) path to be opened by the alarm |

### `attendee`

This class represents a attendee.

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `display name` | `text` | r/o | The first and last name of the attendee. |
| `email` | `text` | r/o | e-mail of the attendee. |
| `participation status` | `participation status` | r/o | The invitation status for the attendee. |

### `event`

This class represents an event.

**Contains:** `attendee`, `display alarm`, `mail alarm`, `open file alarm`, `sound alarm`

**Properties:**

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `description` | `text` | r/w | The events notes. |
| `start date` | `date` | r/w | The event start date. |
| `end date` | `date` | r/w | The event end date. |
| `allday event` | `boolean` | r/w | True if the event is an all-day event |
| `recurrence` | `text` | r/w | The iCalendar (RFC 2445) string describing the event recurrence, if defined |
| `sequence` | `integer` | r/o | The event version. |
| `stamp date` | `date` | r/w | The event modification date. |
| `excluded dates` | `date` | r/w | The exception dates. |
| `status` | `event status` | r/w | The event status. |
| `summary` | `text` | r/w | This is the event summary. |
| `location` | `text` | r/w | This is the event location. |
| `uid` | `text` | r/o | An unique event key. |
| `url` | `text` | r/w | The URL associated to the event. |

**Responds to:** `show`

### Enumerations

### `participation status`

| Value | Description |
|-------|-------------|
| `unknown` | No anwser yet |
| `accepted` | Invitation has been accepted |
| `declined` | Invitation has been declined |
| `tentative` | Invitation has been tentatively accepted |

### `event status`

| Value | Description |
|-------|-------------|
| `cancelled` | A cancelled event |
| `confirmed` | A confirmed event |
| `none` | An event without status |
| `tentative` | A tentative event |

### `calendar priority`

| Value | Description |
|-------|-------------|
| `no priority` | No priority |
| `low priority` | Low priority |
| `medium priority` | Medium priority |
| `high priority` | High priority |

### `view type`

| Value | Description |
|-------|-------------|
| `day view` | The iCal day view |
| `week view` | The iCal week view |
| `month view` | The iCal month view |
