"""The floating dot: drag/click to summon the composer; hosts workers, popups, guide overlay."""
import os
import sys
import math
import time
import threading

from PySide6 import QtCore, QtGui, QtWidgets

import agent
from axon import config
from axon.settings import load_settings, save_settings
from axon.ui.theme import (ACCENT, ACCENT_2, PANEL_BG, PANEL_BG_2, SURFACE, BORDER, TEXT,
                            MUTED, FONT_FAMILY, FONT_CSS, HEADER_ICON_SIZE, CONTROL_ICON_SIZE)
from axon.ui.widgets import (IconButton, ArrowButton, ActivityButton, CloseButton,
    ProfileButton, GuideButton, ModeButton, ApprovalButton, CameraButton)
from axon.ui.composer import Composer
from axon.ui.popups import ApprovalPopup, FolderPickPopup
from axon.ui.highlight import HighlightOverlay, PanelFrame, MarkView
from axon.workers import AgentWorker, InboxWatcher, GuideWorker


class FloatingDot(QtWidgets.QWidget):
    DIAM = 44
    inbox_suggestions = QtCore.Signal(object)  # folder suggestions ready (from a worker thread)
    hotkey_email = QtCore.Signal(str, str, str)  # (eid, subject, sender) from the global hotkey
    proactive_nudge = QtCore.Signal(object)      # an upcoming event to gently offer help with
    browser_setup_msg = QtCore.Signal(str)       # result of the one-time browser sign-in
    followup_due = QtCore.Signal(object)         # a follow-up reminder (no reply yet) from the add-in
    _HOTKEY_ID = 0xA17  # Ctrl+Alt+M -> file the email currently open/selected in Outlook

    def __init__(self):
        super().__init__()
        self.setWindowFlags(
            QtCore.Qt.FramelessWindowHint
            | QtCore.Qt.WindowStaysOnTopHint
            | QtCore.Qt.Tool
        )
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setFixedSize(self.DIAM, self.DIAM)
        self.setCursor(QtCore.Qt.PointingHandCursor)

        self._drag_offset = None
        self._moved = False
        self._busy = False
        self._hover = False
        self._spin_start = time.monotonic()
        self._spin_angle = 0.0
        self._pulse = 0.0
        self._spin_timer = QtCore.QTimer(self)
        self._spin_timer.setInterval(16)
        self._spin_timer.timeout.connect(self._advance_spin)

        self._chat_history = []  # 'Ask Maia' conversation memory (text only), kept across turns
        self.composer = Composer()
        self.composer.submitted.connect(self._start_task)
        self.composer.guide_requested.connect(self._start_guide)
        self.composer.ask_requested.connect(lambda q: self._start_task(q, mode="chat"))
        self.composer.cancel_requested.connect(self._cancel_task)
        self.composer.dismissed.connect(self._show_dot)

        self.highlight = HighlightOverlay()
        self._approval_popup = ApprovalPopup()
        self._approval_popup.decided.connect(self._on_approval_decided)

        # ---- Inbox auto-filer (toggled in the profile dropdown menu) ----
        self._settings = load_settings()
        _act = self.composer.profile_btn.autofile_action
        _act.setChecked(self._settings.get("autofile", False))  # set before connecting, so it won't fire
        _act.toggled.connect(self._set_autofile)
        self.composer.profile_btn.signin_action.triggered.connect(self._setup_browser)
        self.browser_setup_msg.connect(self.composer.append_log)
        self._inbox_watcher = InboxWatcher()
        self._inbox_watcher.opened.connect(self._on_inbox_opened)
        self._folder_popup = FolderPickPopup()
        self._folder_popup.chosen.connect(self._on_folder_chosen)
        self.inbox_suggestions.connect(self._show_folder_popup)
        # Global hotkey (Ctrl+Alt+M): file whatever email is open/selected in Outlook right now.
        self.hotkey_email.connect(self._on_inbox_opened)
        self._register_hotkey()
        if self._settings.get("autofile", False):
            self._inbox_watcher.start()
        _app = QtWidgets.QApplication.instance()
        if _app is not None:
            _app.aboutToQuit.connect(self._inbox_watcher.stop)

        self._init_tray()   # system-tray icon: completion notifications + show/quit menu

        # Proactive nudges: shortly before a meeting, offer to prep. Checked on a timer.
        self._pending_suggestion = None
        self._proactive_seen = set()
        self.proactive_nudge.connect(self._on_proactive)
        self._proactive_timer = QtCore.QTimer(self)
        self._proactive_timer.setInterval(5 * 60 * 1000)   # every 5 minutes
        self._proactive_timer.timeout.connect(self._proactive_check)
        # Follow-up reminders recorded by the Outlook add-in.
        self.followup_due.connect(self._on_followup_due)
        self._followup_timer = QtCore.QTimer(self)
        self._followup_timer.setInterval(2 * 60 * 1000)    # every 2 minutes
        self._followup_timer.timeout.connect(self._followup_check)
        if config.IS_WINDOWS:
            self._proactive_timer.start()
            self._followup_timer.start()
            QtCore.QTimer.singleShot(60 * 1000, self._proactive_check)  # first check ~1 min after launch
            QtCore.QTimer.singleShot(30 * 1000, self._followup_check)

        # Tell the backend which window is the dot, so opened apps land on the dot's monitor.
        try:
            config.DOT_HWND = int(self.winId())
            config.DOT_PID = os.getpid()
        except Exception:
            pass

        screen = QtWidgets.QApplication.primaryScreen().availableGeometry()
        self.move(screen.right() - self.DIAM - 30, screen.bottom() - self.DIAM - 80)
        QtCore.QTimer.singleShot(0, self.composer.prepare_for_open)

        self._thread = None
        self._worker = None
        self._guide_thread = None
        self._guide_worker = None

    def paintEvent(self, _):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)

        glow = QtCore.QRectF(4, 4, self.DIAM - 8, self.DIAM - 8)
        glow_grad = QtGui.QRadialGradient(glow.center(), glow.width() / 2)
        glow_alpha = 62 + int(34 * self._pulse) if self._busy else (54 if self._hover else 32)
        glow_grad.setColorAt(0, QtGui.QColor(255, 255, 255, glow_alpha))
        glow_grad.setColorAt(1, QtGui.QColor(255, 255, 255, 0))
        p.setBrush(QtGui.QBrush(glow_grad))
        p.setPen(QtCore.Qt.NoPen)
        p.drawEllipse(glow)

        rect = QtCore.QRectF(6, 6, self.DIAM - 12, self.DIAM - 12)
        fill_alpha = 92 + int(20 * self._pulse) if self._busy else (104 if self._hover else 72)
        p.setBrush(QtGui.QColor(18, 19, 25, fill_alpha))
        ring = QtGui.QColor("#ffb45c") if self._busy else QtGui.QColor(255, 255, 255, 72)
        p.setPen(QtGui.QPen(ring, 1.4 if self._busy else 1.2))
        p.drawEllipse(rect)

        star = QtGui.QPainterPath()
        cx, cy = self.DIAM / 2, self.DIAM / 2
        radius = 12.4 + (1.2 * self._pulse if self._busy else 0)
        star.moveTo(0, -radius)
        star.cubicTo(2.8, -5.2, 5.2, -2.8, radius, 0)
        star.cubicTo(5.2, 2.8, 2.8, 5.2, 0, radius)
        star.cubicTo(-2.8, 5.2, -5.2, 2.8, -radius, 0)
        star.cubicTo(-5.2, -2.8, -2.8, -5.2, 0, -radius)
        transform = QtGui.QTransform()
        transform.translate(cx, cy)
        if self._busy:
            transform.rotate(self._spin_angle)
        star = transform.map(star)
        p.setBrush(QtGui.QColor("#ffffff"))
        p.setPen(QtCore.Qt.NoPen)
        p.drawPath(star)

        if self._busy:
            dot = QtCore.QRectF(cx + 10, cy - 13, 5, 5)
            p.setBrush(QtGui.QColor("#ffd166"))
            p.drawEllipse(dot)

    def _advance_spin(self):
        elapsed = time.monotonic() - self._spin_start
        self._spin_angle = (elapsed * 150) % 360
        self._pulse = (1 + math.sin(elapsed * 5.2)) / 2
        self.update()

    def _set_busy(self, busy):
        if self._busy == busy:
            return
        self._busy = busy
        if busy:
            self._spin_start = time.monotonic()
            self._spin_timer.start()
        else:
            self._spin_timer.stop()
            self._spin_angle = 0.0
            self._pulse = 0.0
        self.update()

    def enterEvent(self, _):
        self._hover = True
        self.update()

    def leaveEvent(self, _):
        self._hover = False
        self.update()

    def mousePressEvent(self, e):
        if e.button() == QtCore.Qt.LeftButton:
            self._drag_offset = e.globalPosition().toPoint() - self.frameGeometry().topLeft()
            self._moved = False

    def mouseMoveEvent(self, e):
        if self._drag_offset is not None:
            self.move(e.globalPosition().toPoint() - self._drag_offset)
            self._moved = True

    def mouseReleaseEvent(self, e):
        if e.button() == QtCore.Qt.LeftButton:
            if not self._moved:
                self._toggle_composer()
            self._drag_offset = None

    def _tray_icon(self):
        pix = QtGui.QPixmap(32, 32)
        pix.fill(QtCore.Qt.transparent)
        p = QtGui.QPainter(pix)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        grad = QtGui.QLinearGradient(0, 0, 32, 32)
        grad.setColorAt(0, QtGui.QColor(ACCENT))
        grad.setColorAt(1, QtGui.QColor(ACCENT_2))
        p.setBrush(QtGui.QBrush(grad))
        p.setPen(QtCore.Qt.NoPen)
        p.drawEllipse(4, 4, 24, 24)
        p.end()
        return QtGui.QIcon(pix)

    def _init_tray(self):
        """System-tray presence: shows task-done notifications and a show/quit menu."""
        self._tray = None
        try:
            if not QtWidgets.QSystemTrayIcon.isSystemTrayAvailable():
                return
            self._tray = QtWidgets.QSystemTrayIcon(self._tray_icon(), self)
            self._tray.setToolTip("Axon intelligence")
            menu = QtWidgets.QMenu()
            menu.addAction("Show Axon", self._summon)
            menu.addAction("Quit", QtWidgets.QApplication.instance().quit)
            self._tray.setContextMenu(menu)
            self._tray.activated.connect(
                lambda r: self._summon() if r == QtWidgets.QSystemTrayIcon.Trigger else None)
            self._tray.messageClicked.connect(self._summon)
            self._tray.show()
        except Exception:
            self._tray = None

    def _summon(self):
        """Bring up the composer (from the tray icon or a notification click)."""
        if not self.composer.isVisible():
            self._toggle_composer()
        else:
            self.composer.raise_()
            self.composer.activateWindow()
        if self._pending_suggestion:   # a proactive nudge was clicked — prefill it
            try:
                self.composer.input.setPlainText(self._pending_suggestion)
                self.composer.input.setFocus()
            except Exception:
                pass
            self._pending_suggestion = None

    def _setup_browser(self):
        """One-time browser sign-in from the profile menu (runs off the UI thread)."""
        if not self.composer._activity_visible:
            self.composer.activity_btn.setChecked(True)
        self.composer.append_log("Opening Axon's browser — sign in to the accounts you want Axon to use…")

        def work():
            try:
                res = agent.setup_browser({})
                text = res["content"][0]["text"]
            except Exception as e:
                text = f"Couldn't open the browser: {e}"
            self.browser_setup_msg.emit(text)
        threading.Thread(target=work, daemon=True).start()

    def _followup_check(self):
        """Poll the add-in's follow-up file; remind (tray) about any now due with no reply yet."""
        def work():
            try:
                for it in (agent.due_followups() or []):
                    self.followup_due.emit(it)
            except Exception:
                pass
        threading.Thread(target=work, daemon=True).start()

    @QtCore.Slot(object)
    def _on_followup_due(self, it):
        who = (it.get("who") or "someone").split("<")[0].strip()
        subject = it.get("subject") or "your email"
        self._pending_suggestion = (
            f"Draft a follow-up email to {who} about \"{subject}\" — they haven't replied yet.")
        if self._tray:
            try:
                self._tray.showMessage(
                    "Follow-up due — no reply yet",
                    f"{who} hasn't replied about “{subject}”. Click to draft a follow-up.",
                    self._tray_icon(), 10000)
            except Exception:
                pass

    def _proactive_check(self):
        """Background-poll Outlook for an imminent meeting; emit a nudge if there's a new one."""
        def work():
            try:
                ev = agent.upcoming_event(10)
            except Exception:
                ev = None
            if ev and ev.get("id") and ev["id"] not in self._proactive_seen:
                self._proactive_seen.add(ev["id"])
                self.proactive_nudge.emit(ev)
        threading.Thread(target=work, daemon=True).start()

    @QtCore.Slot(object)
    def _on_proactive(self, ev):
        subj = ev.get("subject") or "a meeting"
        mins = ev.get("minutes", 0)
        self._pending_suggestion = f"Prep me for my meeting \"{subj}\": pull recent emails and context about it."
        when = "now" if mins <= 0 else f"in {mins} min"
        if self._tray:
            try:
                self._tray.showMessage(
                    "Upcoming meeting", f"“{subj}” {when} — click to have Axon prep you.",
                    self._tray_icon(), 9000)
            except Exception:
                pass

    def _notify_done(self, result):
        """Toast the result if the user isn't looking at the panel, so they can walk away while it works."""
        try:
            if not self._tray:
                return
            if self.composer.isVisible() and self.composer.isActiveWindow():
                return
            summary = " ".join(str(result).split())
            if len(summary) > 200:
                summary = summary[:197] + "..."
            self._tray.showMessage("Axon — task done", summary or "Done.", self._tray_icon(), 7000)
        except Exception:
            pass

    def _toggle_composer(self):
        dot = self.frameGeometry()
        cw, ch = self.composer.width(), self.composer.height()
        screen = self._screen_for_dot().availableGeometry()
        margin = 10
        x = min(dot.left(), screen.right() - cw - margin)
        y = dot.top() - ch - 10
        if y < screen.top() + margin:
            y = dot.bottom() + margin
        x = max(screen.left() + margin, min(x, screen.right() - cw - margin))
        y = max(screen.top() + margin, min(y, screen.bottom() - ch - margin))
        self.composer.move(x, y)
        self.composer.show()
        self.composer.raise_()
        self.composer.input.setFocus()
        self.hide()

    def _screen_for_dot(self):
        center = self.frameGeometry().center()
        return (
            QtWidgets.QApplication.screenAt(center)
            or self.screen()
            or QtWidgets.QApplication.primaryScreen()
        )

    def _show_dot(self):
        self.show()
        self.raise_()
        self.update()

    def _is_running(self):
        return bool(self._thread or self._guide_thread)

    def _set_active_monitor(self):
        """Use the COMPOSER's screen (where the user just typed) as the reference for placing apps,
        falling back to the dot. Call while the composer is still visible (before dismiss)."""
        try:
            ref = self.composer if self.composer.isVisible() else self
            config.DOT_HWND = int(ref.winId())
            config.DOT_PID = os.getpid()
        except Exception:
            pass

    def _start_task(self, question, mode="agent"):
        if self._busy or self._is_running():
            self.composer.task_finished()
            return
        self._set_active_monitor()
        image_path = self.composer.attached_image()  # screenshot, if the user attached one
        self.composer.task_started()
        self._set_busy(True)
        self.composer.clear_attachment(delete=False)  # the worker owns the file until it finishes
        self.composer.dismiss()

        self._thread = QtCore.QThread()
        self._worker = AgentWorker(
            question, approval_mode=self.composer.approve_btn.ask_before(),
            image_path=image_path, mode=mode,
            chat_history=list(self._chat_history) if mode == "chat" else None)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.status.connect(self.composer.append_log)
        self._worker.plan.connect(self.composer.set_plan)
        self._worker.approval_requested.connect(self._on_approval_requested)
        self._worker.finished.connect(self._on_finished)
        self._thread.start()

    def _active_screen(self):
        try:
            return self.composer.screen() or self.screen()
        except Exception:
            return self.screen()

    @QtCore.Slot(str)
    def _on_approval_requested(self, description):
        self._approval_popup.ask(description, self._active_screen())  # where the composer/dot is

    @QtCore.Slot(bool)
    def _on_approval_decided(self, approved):
        if self._worker:
            self._worker.resolve_approval(approved)

    # ---- profile / inbox auto-filer ----
    def _set_autofile(self, on):
        self._settings["autofile"] = bool(on)
        save_settings(self._settings)
        if on:
            self._inbox_watcher.start()
        else:
            self._inbox_watcher.stop()

    def _register_hotkey(self):
        """Register a system-wide Ctrl+Alt+M hotkey (Windows). WM_HOTKEY is delivered to this
        window and handled in nativeEvent(). The OS releases it automatically on exit."""
        if sys.platform != "win32":
            return
        try:
            import ctypes
            MOD_ALT, MOD_CONTROL, MOD_NOREPEAT, VK_M = 0x0001, 0x0002, 0x4000, 0x4D
            ctypes.windll.user32.RegisterHotKey(
                int(self.winId()), self._HOTKEY_ID, MOD_CONTROL | MOD_ALT | MOD_NOREPEAT, VK_M)
        except Exception:
            pass

    def nativeEvent(self, eventType, message):
        if eventType == b"windows_generic_MSG":
            try:
                import ctypes
                from ctypes import wintypes
                msg = wintypes.MSG.from_address(int(message))
                if msg.message == 0x0312 and msg.wParam == self._HOTKEY_ID:  # WM_HOTKEY
                    self._hotkey_fired()
            except Exception:
                pass
        return super().nativeEvent(eventType, message)

    def _hotkey_fired(self):
        """On the hotkey: fetch the selected/open Outlook email (off the UI thread) and file it."""
        def work():
            try:
                info = agent.active_email()
            except Exception:
                info = None
            if info:
                self.hotkey_email.emit(info[0], info[1], info[2])
        threading.Thread(target=work, daemon=True).start()

    def _on_inbox_opened(self, eid, subject, sender):
        # Suggest folders off the UI thread (it reads Outlook + calls the model), then show the popup.
        def work():
            try:
                data = agent.suggest_folders(eid)
            except Exception:
                data = None
            if data:
                data["eid"] = eid
                data["subject"] = data.get("subject") or subject
                data["sender"] = data.get("sender") or sender
                self.inbox_suggestions.emit(data)
        threading.Thread(target=work, daemon=True).start()

    @QtCore.Slot(object)
    def _show_folder_popup(self, data):
        if not data or not data.get("folders"):
            return  # nothing to file into
        self._folder_popup.present(
            data["eid"], data.get("subject", ""), data.get("sender", ""),
            data.get("suggestions", []), data.get("folders", []), self._active_screen())

    def _on_folder_chosen(self, eid, folder):
        def work():
            try:
                agent.move_email_to_folder(eid, folder)
            except Exception:
                pass
        threading.Thread(target=work, daemon=True).start()

    @QtCore.Slot(str)
    def _on_finished(self, result):
        self._approval_popup.hide()
        w = self._worker
        if w and getattr(w, "mode", "") == "chat":
            self.composer.append_answer_label()   # label the reply in the conversation thread
        else:
            self.composer.append_log("")
        self.composer.append_log(result)
        self.composer.task_finished()
        if not (w and getattr(w, "mode", "") == "chat"):
            self._notify_done(result)   # background completion toast for agent tasks
        if w and getattr(w, "mode", "") == "chat" and not str(result).startswith("Error:"):
            # Remember this Maia exchange so follow-up questions keep context (cap to recent turns).
            self._chat_history.append({"role": "user", "content": (w.question or "").strip() or "[screenshot]"})
            self._chat_history.append({"role": "assistant", "content": str(result)})
            self._chat_history = self._chat_history[-12:]
        if w and getattr(w, "image_path", None):
            try:
                os.remove(w.image_path)  # clean up the attached screenshot temp file
            except Exception:
                pass
        if self._thread:
            self._thread.quit()
            self._thread.wait()
            self._thread = None
            self._worker = None
        self._set_busy(False)

    def _start_guide(self, question):
        if self._busy or self._is_running():
            self.composer.task_finished()
            return
        self._set_active_monitor()
        self.composer.task_started()
        self._set_busy(True)
        self.composer.dismiss()

        self.composer.append_log("Guiding you step by step — do each step and I'll point to the next.")
        self._guide_thread = QtCore.QThread()
        self._guide_worker = GuideWorker(question)
        self._guide_worker.moveToThread(self._guide_thread)
        self._guide_thread.started.connect(self._guide_worker.run)
        self._guide_worker.step.connect(self._on_guide_step)
        self._guide_worker.finished.connect(self._on_guide_finished)
        self._guide_thread.start()

    @QtCore.Slot(str, object, bool)
    def _on_guide_step(self, instruction, marker, done):
        if instruction:
            self.composer.append_log(("✓ " if done else "→ ") + instruction)
        try:
            scr = self._active_screen()  # composer/dot monitor — matches the cropped guide screenshot
            if isinstance(marker, dict) and marker.get("type") == "clickhere":
                self.highlight.show_click(
                    marker["fx"], marker["fy"], marker.get("label", "Click here"),
                    persist=not done, screen=scr)
            elif isinstance(marker, dict) and marker.get("type") == "box":
                self.highlight.show_box(
                    marker["fx"], marker["fy"], marker["fw"], marker["fh"],
                    marker.get("label", ""), persist=not done, screen=scr)
            elif isinstance(marker, dict) and marker.get("type") == "point":
                self.highlight.show_pointer(
                    marker["fx"], marker["fy"], marker.get("label", ""),
                    persist=not done, hide_after_ms=6000, screen=scr)
            elif done:
                self.highlight.hide_after(2500)
        except Exception:
            pass

    @QtCore.Slot()
    def _on_guide_finished(self):
        try:
            self.highlight.hide_after(6000)
        except Exception:
            pass
        self.composer.task_finished()
        if self._guide_thread:
            self._guide_thread.quit()
            self._guide_thread.wait()
            self._guide_thread = None
            self._guide_worker = None
        self._set_busy(False)

    def _cancel_task(self):
        if not self._is_running():
            self.composer.task_finished()
            return
        self._approval_popup.hide()
        self.composer.append_log("")
        self.composer.append_log("Stopping...")
        self.composer.task_finished()
        self.composer.dismiss()
        self._set_busy(False)
        if self._worker:
            self._worker.cancel()
        if self._thread:
            self._thread.requestInterruption()
        if self._guide_worker:
            self._guide_worker.cancel()
        if self._guide_thread:
            self._guide_thread.requestInterruption()
