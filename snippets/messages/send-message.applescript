-- title: Send iMessage to a contact
-- verified: no
-- macos: 13+
-- app: Messages
-- description: Send an iMessage to a contact by phone number or email
-- source: app dictionary
-- notes: Recipient must be reachable via iMessage. Use phone number or Apple ID email.

set recipientID to "+1234567890" -- phone number or email
set messageText to "Hello from AppleScript!"

tell application "Messages"
	-- Find the iMessage service
	set targetService to 1st account whose service type = iMessage
	-- Get or create a chat with the recipient
	set targetChat to make new text chat with properties {participants:{participant recipientID of targetService}}
	send messageText to targetChat
end tell
