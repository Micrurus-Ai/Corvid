# Axon intelligence

Axon intelligence is a local desktop assistant with a floating star launcher and a movable composer. It uses OpenAI plus local computer-control tools to inspect the screen, open apps, and perform desktop tasks.

## Run

```powershell
cd assistant
.\run.bat
```

Or run it directly:

```powershell
cd assistant
.\.venv\Scripts\python.exe overlay.py
```

## Setup

Create `assistant/.env` with:

```text
OPENAI_API_KEY=sk-...
```

The assistant expects the local `open-computer-use` command to be available on `PATH`.

## Controls

- Click the floating star to open the composer.
- Drag the star to move it between screens.
- Press `Ctrl+Enter` or the arrow button to send.
- While running, the composer closes back to the rotating star.
- Reopen the composer during a run to see activity or press the stop square.
- Use `Esc` or the `X` button to close the composer.
