# Bluetooth file exchange AppleScript Dictionary

> Auto-generated from `Bluetooth File Exchange.sdef` inside the app bundle.  
> macOS 15.6. Do not edit manually — regenerate with `uv run tools/sdef_to_md.py "Bluetooth file exchange"`

## Table of Contents

- [Bluetooth File Exchange Suite](#bluetooth-file-exchange-suite)

---

## Bluetooth File Exchange Suite

The suite of commands for Bluetooth File Exchange

### Commands

### `browse`

Browse a device.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `device` | `text` | yes | The device to browse. |

### `send`

Send a file to a bluetooth device.

**Parameters:**

| Parameter | Type | Optional | Description |
|-----------|------|----------|-------------|
| `file` | `file` | yes | The file to send. |
| `to device` | `text` | yes | The device to send the file to. |
