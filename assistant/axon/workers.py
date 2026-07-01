"""Background QObject workers (run on QThreads): the agent loop, the inbox watcher, and the
live guide. They call into the agent backend and emit Qt signals back to the UI."""
import os
import time
import threading

from PySide6 import QtCore

import agent


class AgentWorker(QtCore.QObject):
    status = QtCore.Signal(str)
    finished = QtCore.Signal(str)
    plan = QtCore.Signal(list)               # live checklist: [{task, done}, ...]
    approval_requested = QtCore.Signal(str)  # describes the pending action; UI must call resolve_approval

    def __init__(self, question, approval_mode=False, image_path=None, mode="agent", chat_history=None):
        super().__init__()
        self.question = question
        self.approval_mode = approval_mode
        self.image_path = image_path
        self.mode = mode
        self.chat_history = chat_history or []
        self._cancel_requested = False
        self._approval_event = threading.Event()
        self._approval_result = False

    def cancel(self):
        self._cancel_requested = True
        self._approval_event.set()  # unblock any pending approval

    def _should_cancel(self):
        thread = QtCore.QThread.currentThread()
        return self._cancel_requested or thread.isInterruptionRequested()

    def resolve_approval(self, approved):
        """Called from the UI thread when the user clicks Approve/Skip."""
        self._approval_result = bool(approved)
        self._approval_event.set()

    def _on_approval(self, description):
        """Runs in the worker thread: block until the user approves/skips the action."""
        if not self.approval_mode:
            return True
        self._approval_result = False
        self._approval_event.clear()
        self.approval_requested.emit(description)
        while not self._approval_event.wait(0.15):
            if self._should_cancel():
                return False
        return self._approval_result

    @QtCore.Slot()
    def run(self):
        try:
            if self.mode == "chat":
                result = agent.chat(
                    self.question,
                    on_status=lambda s: self.status.emit(s),
                    image_path=self.image_path,
                    history=self.chat_history,
                )
            else:
                result = agent.run_task(
                    self.question,
                    on_status=lambda s: self.status.emit(s),
                    should_cancel=self._should_cancel,
                    on_approval=self._on_approval,
                    image_path=self.image_path,
                    on_plan=lambda t: self.plan.emit(t),
                )
            self.finished.emit(result or "Done.")
        except Exception as e:
            self.finished.emit(f"Error: {e}")


class InboxWatcher(QtCore.QObject):
    """Runs agent's background Outlook watcher and emits `opened` when an unfiled Inbox email is
    opened/read, so the UI can offer to file it. Toggled on/off from the profile panel."""
    opened = QtCore.Signal(str, str, str)  # entry_id, subject, sender

    def __init__(self):
        super().__init__()
        self._proc = None
        self._thread = None
        self._stop = False

    def is_running(self):
        return self._proc is not None

    def start(self):
        if self._proc is not None:
            return
        self._stop = False
        self._proc = agent.inbox_watcher_popen()
        if self._proc is None:
            return
        self._thread = threading.Thread(target=self._read, daemon=True)
        self._thread.start()

    def _read(self):
        try:
            for line in self._proc.stdout:
                if self._stop:
                    break
                line = (line or "").strip()
                if line.startswith("OPENED|"):
                    parts = line.split("|", 3)
                    if len(parts) >= 4:
                        self.opened.emit(parts[1], parts[2], parts[3])
        except Exception:
            pass

    def stop(self):
        self._stop = True
        if self._proc is not None:
            try:
                self._proc.terminate()
            except Exception:
                pass
            self._proc = None


class GuideWorker(QtCore.QObject):
    """Live coaching: points at the next step, waits for the user to act, then advances.
    Emits one `step` per action; takes no actions itself."""
    step = QtCore.Signal(str, object, bool)  # instruction, pointer dict or None, done
    finished = QtCore.Signal()

    def __init__(self, question):
        super().__init__()
        self.question = question
        self._cancel_requested = False

    def cancel(self):
        self._cancel_requested = True

    def _should_cancel(self):
        thread = QtCore.QThread.currentThread()
        return self._cancel_requested or thread.isInterruptionRequested()

    @QtCore.Slot()
    def run(self):
        try:
            agent.guide_live(
                self.question,
                on_step=lambda p: self.step.emit(
                    p.get("instruction", ""), p.get("marker"), bool(p.get("done"))),
                should_cancel=self._should_cancel,
            )
        except Exception as e:
            self.step.emit(f"Error: {e}", None, True)
        finally:
            self.finished.emit()
