-- title: Open a URL in Safari
-- verified: yes
-- macos: 13+
-- app: Safari
-- description: Opens a URL in a new tab in the front Safari window

set targetURL to "https://example.com"

tell application "Safari"
    activate
    tell front window
        set newTab to make new tab
        set URL of newTab to targetURL
        set current tab to newTab
    end tell
end tell
