-- title: List recent chat participants
-- verified: no
-- macos: 13+
-- app: Messages
-- description: Get a list of recent chat names/participants from Messages
-- source: app dictionary

tell application "Messages"
	set chatList to {}
	repeat with c in chats
		set chatName to name of c
		set chatPeople to {}
		repeat with p in participants of c
			set end of chatPeople to full name of p
		end repeat
		set end of chatList to {chatName, chatPeople}
	end repeat
	return chatList
end tell
