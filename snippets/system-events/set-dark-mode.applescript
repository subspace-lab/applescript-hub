-- title: Toggle dark mode
-- verified: no
-- macos: 13+
-- app: System Events
-- description: Toggle macOS dark mode on or off
-- source: original

tell application "System Events"
	tell appearance preferences
		set dark mode to not dark mode
	end tell
end tell
