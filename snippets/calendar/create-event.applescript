-- title: Create a new calendar event
-- verified: yes
-- macos: 13+
-- app: Calendar
-- description: Creates a new event in a named calendar
-- notes: Change calendarName to match one of your calendars

set calendarName to "Home"
set eventTitle to "AppleScript Test Event"
set startDate to current date
set endDate to startDate + (1 * hours)

tell application "Calendar"
    tell calendar calendarName
        make new event with properties {
            summary: eventTitle,
            start date: startDate,
            end date: endDate,
            description: "Created by AppleScript"
        }
    end tell
end tell
