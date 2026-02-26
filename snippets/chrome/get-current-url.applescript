-- title: Get URL of active tab in Chrome
-- verified: no
-- macos: 13+
-- app: Google Chrome
-- description: Get the URL and title of the active tab in the frontmost Chrome window
-- source: app dictionary

tell application "Google Chrome"
	set tabURL to URL of active tab of front window
	set tabTitle to title of active tab of front window
	return {tabURL, tabTitle}
end tell
