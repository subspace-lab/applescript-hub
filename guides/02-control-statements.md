# Control Statements

## if / else if / else

```applescript
if ageOfCat > 1 then display dialog "Not a kitten."

-- Multi-branch:
if currentTemp < 60 then
    set response to "Chilly"
else if currentTemp > 80 then
    set response to "Hot"
else
    set response to "Nice"
end if
```

## repeat

```applescript
-- Forever (use exit repeat to break):
repeat
    if done then exit repeat
end repeat

-- N times:
repeat 5 times
    -- ...
end repeat

-- Numeric range:
repeat with n from 1 to 10
    log n
end repeat

-- Over a list:
repeat with item in {"a", "b", "c"}
    log item
end repeat

-- While / Until:
repeat while notDone
    set notDone to doWork()
end repeat

repeat until finished
    set finished to doWork()
end repeat
```

## try / on error

```applescript
try
    word 5 of "one two three"  -- triggers error
on error errMsg number errNum
    log "Error " & errNum & ": " & errMsg
end try

-- Handle specific error code only:
try
    display alert "Hello" buttons {"Cancel", "OK"} cancel button 1
on error number -128
    -- user canceled, ignore
end try

-- Re-raise after handling:
try
    riskyOperation()
on error m number n
    cleanUp()
    error m number n  -- re-raise
end try
```

## tell

```applescript
-- Simple:
tell front window of application "Finder" to close

-- Compound (multiple commands to same target):
tell application "Finder"
    close front window
    empty trash
end tell

-- Nested:
tell application "Finder"
    tell front window
        get name of every item
    end tell
end tell
```

## with timeout

Default Apple Event timeout is 2 minutes. Override it:

```applescript
tell application "TextEdit"
    with timeout of 30 seconds
        close document 1 saving ask
    end timeout
end tell

-- Disable timeout entirely:
with timeout of 0 seconds
    -- ...
end timeout
```

## considering / ignoring

Control how text comparisons and app responses behave:

```applescript
-- Text attributes:
ignoring white space
    "Hello Bob" = "HelloBob"  -- true
end ignoring

considering case
    "BOB" = "bob"  -- false
end considering

ignoring case and diacriticals
    "Babs" = "bábs"  -- true
end ignoring

-- Application responses (fire-and-forget):
tell application "Finder"
    ignoring application responses
        empty the trash  -- don't wait
    end ignoring
end tell
```

## use (version / framework declarations)

```applescript
-- Require minimum AppleScript version:
use AppleScript version "2.4"

-- Declare app dependency (enables unqualified terms):
use Safari : application "Safari"
get the name of Safari's front window

-- Load scripting additions:
use scripting additions

-- Load Objective-C frameworks (ASObjC):
use framework "Foundation"
use framework "AppKit"
```

## error (signal errors)

```applescript
error "Something went wrong"
error "Invalid input" number 1001
```
