# do shell script

Run Unix shell commands from AppleScript. Uses `/bin/sh`.

Source: [Apple TN2065](https://developer.apple.com/library/archive/technotes/tn2065/_index.html)

## Basic Usage

```applescript
do shell script "echo hello"
-- result: "hello"

-- Capture output into variable:
set today to do shell script "date +%Y-%m-%d"
```

## Passing Variables — Always Use `quoted form of`

**Never** concatenate variables directly. Use `quoted form of` to safely escape spaces and special characters:

```applescript
-- WRONG (breaks with spaces, $, *, etc.):
do shell script "ls " & myPath

-- CORRECT:
do shell script "ls " & quoted form of myPath
```

```applescript
-- Pipeline example:
set s to "hello world"
do shell script "echo " & quoted form of s & " | tr a-z A-Z"
-- result: "HELLO WORLD"
```

## File Paths

```applescript
-- POSIX path of an alias:
set f to choose file
do shell script "ls -lh " & quoted form of POSIX path of f

-- HFS → POSIX:
POSIX path of alias "Macintosh HD:Users:me:file.txt"
-- result: "/Users/me/file.txt"
```

## Error Handling

Non-zero exit codes raise an error. Stderr becomes the error message.

```applescript
try
    do shell script "ls /nonexistent"
on error errMsg number errNum
    log "Exit " & errNum & ": " & errMsg
end try

-- Capture both stdout and stderr:
do shell script "command 2>&1"
```

## Administrator Privileges

```applescript
do shell script "chmod 755 /usr/local/bin/mytool" with administrator privileges
-- macOS shows an auth dialog
```

## Background Processes

```applescript
-- Fire and forget (don't wait for result):
do shell script "open -a 'Activity Monitor' &> /dev/null &"

-- Get PID of background process:
set pid to do shell script "my_daemon &> /dev/null & echo $!"
-- Later:
do shell script "kill " & pid
```

## State Does Not Persist Between Calls

Each `do shell script` is a new shell process:

```applescript
do shell script "cd ~/Documents"
do shell script "pwd"  -- result: "/" — NOT ~/Documents

-- Chain in one call instead:
do shell script "cd ~/Documents && pwd"
-- result: "/Users/me/Documents"
```

## Line Endings

Output uses carriage returns (`\r`) by default. To preserve Unix `\n`:

```applescript
do shell script "command" without altering line endings
```

## Using a Different Shell

```applescript
do shell script "/bin/bash -c 'echo $BASH_VERSION'"
do shell script "/usr/bin/env python3 -c 'print(42)'"
```

## Common Pitfalls

| Pitfall | Fix |
|---------|-----|
| Path with spaces fails | Use `quoted form of POSIX path of f` |
| `cd` in one call, use dir in next | Chain with `&&` in a single call |
| Interactive command hangs | Use non-interactive flags (e.g. `top -l1`) |
| `~` in path not expanded | Use full paths or `$HOME` in shell |
| Command not found | Use full path: `/usr/bin/python3` not `python3` |
