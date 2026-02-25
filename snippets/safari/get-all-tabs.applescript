-- title: Get all open tab URLs in Safari
-- verified: yes
-- macos: 13+
-- app: Safari
-- description: Returns a list of URLs for every open tab across all Safari windows

tell application "Safari"
    set allURLs to {}
    repeat with aWindow in every window
        repeat with aTab in every tab of aWindow
            set end of allURLs to URL of aTab
        end repeat
    end repeat
    return allURLs
end tell
