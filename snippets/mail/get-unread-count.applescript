-- title: Get unread message count per mailbox
-- verified: yes
-- macos: 13+
-- app: Mail
-- description: Returns unread counts for all mailboxes across all accounts

tell application "Mail"
    set result to {}
    repeat with anAccount in every account
        repeat with aMailbox in every mailbox of anAccount
            set unread to unread count of aMailbox
            if unread > 0 then
                set end of result to (name of aMailbox) & ": " & unread
            end if
        end repeat
    end repeat
    return result
end tell
