-- title: Get selected files in Finder
-- verified: yes
-- macos: 13+
-- app: Finder
-- description: Returns POSIX paths of all currently selected files/folders in Finder

tell application "Finder"
    set selectedItems to selection
    set pathList to {}
    repeat with anItem in selectedItems
        set end of pathList to POSIX path of (anItem as alias)
    end repeat
    return pathList
end tell
