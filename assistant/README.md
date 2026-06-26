# Axon intelligence Desktop Assistant

A floating, always-on-top star opens a compact composer for desktop tasks.

## How It Works

- `overlay.py`: floating star, composer UI, activity panel, guide overlay, and running-state controls.
- `agent.py`: OpenAI agent loop, screen inspection, app launch placement, and tool execution.

## Setup

Create `.env` in this folder:

```text
OPENAI_API_KEY=sk-...
```

Make sure `open-computer-use` is installed and available on `PATH`.

## Run

Double-click `run.bat`, or run:

```powershell
.\.venv\Scripts\python.exe overlay.py
```

## Controls

- Drag the floating star to reposition it.
- Click the star to open the composer.
- Press `Ctrl+Enter` or the arrow button to send.
- The star rotates while the agent is active.
- Reopen the composer during a run to view activity or click the stop square.
- Press `Esc` or `X` to close the composer.
