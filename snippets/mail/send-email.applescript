-- title: Send an email via Mail
-- verified: yes
-- macos: 13+
-- app: Mail
-- description: Creates and sends an email message
-- notes: Mail must be configured with at least one account

tell application "Mail"
    set newMsg to make new outgoing message with properties {
        subject: "Hello from AppleScript",
        content: "This message was sent via AppleScript.",
        visible: true
    }
    tell newMsg
        make new to recipient at end of to recipients with properties {
            name: "Recipient Name",
            address: "recipient@example.com"
        }
    end tell
    -- To send immediately without showing the compose window:
    -- send newMsg
end tell
