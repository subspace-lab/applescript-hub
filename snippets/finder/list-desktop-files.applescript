-- title: List all files on the Desktop
-- verified: yes
-- macos: 13+
-- app: Finder
-- description: Returns names of all items on the Desktop

tell application "Finder"
    set desktopItems to every item of desktop
    set nameList to {}
    repeat with anItem in desktopItems
        set end of nameList to name of anItem
    end repeat
    return nameList
end tell
