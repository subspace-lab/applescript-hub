-- title: List all open tabs in Chrome
-- verified: no
-- macos: 13+
-- app: Google Chrome
-- description: List all open tabs across all Chrome windows with their URLs and titles
-- source: github.com/briangonzalez/awesome-applescripts

tell application "Google Chrome"
	set tabList to {}
	repeat with w in windows
		set winName to name of w
		repeat with t in tabs of w
			set end of tabList to {title of t, URL of t}
		end repeat
	end repeat
	return tabList
end tell
