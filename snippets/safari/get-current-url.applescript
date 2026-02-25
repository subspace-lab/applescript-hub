-- title: Get current tab URL in Safari
-- verified: yes
-- macos: 13+
-- app: Safari
-- description: Returns the URL of the active tab in the frontmost Safari window

tell application "Safari"
    get URL of current tab of front window
end tell
