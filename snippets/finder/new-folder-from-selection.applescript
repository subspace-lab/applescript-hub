-- title: Move selected files into a new folder
-- verified: yes
-- macos: 13+
-- app: Finder
-- description: Creates a new folder in the same location as selected files and moves them into it

tell application "Finder"
    set selectedItems to selection as list
    if (count of selectedItems) is 0 then
        display dialog "No files selected." buttons {"OK"} default button 1
        return
    end if
    set targetFolder to container of (item 1 of selectedItems)
    set newFolder to make new folder at targetFolder with properties {name: "New Folder"}
    repeat with anItem in selectedItems
        move anItem to newFolder
    end repeat
    activate
    reveal newFolder
end tell
