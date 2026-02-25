-- title: Set system output volume
-- verified: yes
-- macos: 13+
-- app: System Events
-- description: Set the system output volume (0-100) and mute state

set volume output volume 50
set volume with output muted

-- To unmute:
-- set volume without output muted

-- To get current volume:
-- get volume settings
