-- title: Send keyboard shortcuts
-- verified: no
-- macos: 13+
-- app: System Events
-- description: Send keyboard shortcuts (e.g. Cmd+C, Cmd+Shift+S) to the frontmost app
-- source: macosxautomation.com
-- notes: Requires assistive access enabled

-- Send Cmd+C (copy)
tell application "System Events"
	keystroke "c" using command down
end tell

-- Send Cmd+Shift+S (save as)
-- tell application "System Events"
-- 	keystroke "s" using {command down, shift down}
-- end tell

-- Send Cmd+Option+Esc (force quit dialog)
-- tell application "System Events"
-- 	key code 53 using {command down, option down}
-- end tell
