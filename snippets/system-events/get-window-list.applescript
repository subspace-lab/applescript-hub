-- title: List all windows of frontmost app
-- verified: no
-- macos: 13+
-- app: System Events
-- description: Get names, positions, and sizes of all windows in the frontmost application
-- source: original

tell application "System Events"
	set frontApp to first application process whose frontmost is true
	set appName to name of frontApp
	set windowInfo to {}
	tell frontApp
		repeat with w in windows
			set winName to name of w
			set winPos to position of w
			set winSize to size of w
			set end of windowInfo to {winName, winPos, winSize}
		end repeat
	end tell
	return {appName, windowInfo}
end tell
