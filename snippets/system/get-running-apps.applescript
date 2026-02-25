-- title: Get list of running applications
-- verified: yes
-- macos: 13+
-- app: System Events
-- description: Returns names of all currently running applications

tell application "System Events"
    set runningApps to name of every process whose background only is false
    return runningApps
end tell
