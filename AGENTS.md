# Axon intelligence

This repository is now focused on the local Axon intelligence desktop assistant.

## Working Rules

- Keep changes small, direct, and easy to verify.
- Prefer existing local patterns over new abstractions.
- Do not add release notes or template documentation unless explicitly requested.
- If a change affects user-facing behavior, update the relevant README or inline help in the same task.
- Reply in the same language the user uses; if the user writes in English, reply in English.
- Before pushing, sync with the remote first.

## Key Entry Points

- `assistant/overlay.py`: floating star, composer UI, activity/history panel, and guide overlay.
- `assistant/agent.py`: agent loop, tool calls, app launch behavior, and screen guidance.
- `assistant/README.md`: local run instructions for the desktop assistant.
