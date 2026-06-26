@echo off
REM Opens the agent's dedicated Chrome (separate from your normal Chrome, port 9222,
REM own persistent profile). Log into Google / any sites here ONCE.
REM The agent reuses THIS exact browser for every "browse" task (it checks port 9222
REM first), so your logins are available automatically. You can leave it open.
start "" chrome --remote-debugging-port=9222 --user-data-dir="%~dp0.chrome-debug-profile" "https://accounts.google.com"
