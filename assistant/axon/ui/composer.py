"""The composer panel: input box, mode/approval/camera controls, attachment bar, and log."""
import os
import re
import html as _htmllib

from PySide6 import QtCore, QtGui, QtWidgets


def _md_to_html(text):
    """Render a line/block of lightweight markdown as HTML for the answer log:
    **bold**, # headers, - bullets, `code`. Keeps everything else as plain text."""
    out = []
    for raw in (text or "").split("\n"):
        s = _htmllib.escape(raw)
        head = re.match(r"\s*#{1,6}\s+(.*)", s)
        if head:
            s = "<b>" + head.group(1).strip() + "</b>"
        else:
            bullet = re.match(r"(\s*)[-*]\s+(.*)", s)
            if bullet:
                s = bullet.group(1) + "&#8226; " + bullet.group(2)
        s = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", s)
        s = re.sub(r"`([^`]+?)`", r"<code>\1</code>", s)
        out.append(s)
    return "<br>".join(out)
from axon.ui.theme import (ACCENT, ACCENT_2, PANEL_BG, PANEL_BG_2, SURFACE, BORDER, TEXT,
                            MUTED, FONT_FAMILY, FONT_CSS, HEADER_ICON_SIZE, CONTROL_ICON_SIZE)
from axon.ui.widgets import (IconButton, ArrowButton, ActivityButton, CloseButton,
    ProfileButton, GuideButton, ModeButton, ApprovalButton, CameraButton, MicButton)
from axon.ui.region_selector import RegionSelector
from axon.ui.highlight import PanelFrame, MarkView
from axon.ui.theme import SURFACE_2


class Composer(QtWidgets.QWidget):
    submitted = QtCore.Signal(str)
    voice_text = QtCore.Signal(str)  # transcribed speech -> input box (emitted from a worker thread)
    QUICK_ACTIONS = [
        ("Brief me on my day", "Give me my daily briefing."),
        ("What needs my attention", "Triage my unread inbox — what needs my attention?"),
        ("Make a branded deck", "Create a branded PowerPoint deck about "),
        ("Research a website + email report", "Research this website and email me a one-page report: "),
        ("Ask my documents", "Based on my project documents, "),
        ("Summarize a PDF", "Summarize this PDF: "),
    ]
    guide_requested = QtCore.Signal(str)
    ask_requested = QtCore.Signal(str)   # 'Ask Maia' — straight to the LLM, no actions
    cancel_requested = QtCore.Signal()
    dismissed = QtCore.Signal()

    def __init__(self):
        super().__init__()
        self.setWindowFlags(
            QtCore.Qt.FramelessWindowHint
            | QtCore.Qt.WindowStaysOnTopHint
            | QtCore.Qt.Tool
        )
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setMinimumSize(340, 200)
        self.resize(420, 212)
        self._drag_offset = None
        self._activity_visible = False
        self._height_before_activity = self.height()
        self._prepared_for_open = False

        self.card = PanelFrame(self)
        self.card.setGeometry(8, 8, self.width() - 16, self.height() - 16)

        layout = QtWidgets.QVBoxLayout(self.card)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(12)

        header = QtWidgets.QHBoxLayout()
        header.setSpacing(10)
        mark = MarkView()
        header.addWidget(mark)

        title_block = QtWidgets.QVBoxLayout()
        title_block.setSpacing(1)
        title = QtWidgets.QLabel("Axon intelligence")
        title.setStyleSheet(f"color:{TEXT};{FONT_CSS}font-size:15px;font-weight:700;")
        title_block.addWidget(title)
        header.addLayout(title_block)
        header.addStretch(1)

        self.profile_btn = ProfileButton()  # opens a small dropdown menu (auto-file toggle, etc.)
        header.addWidget(self.profile_btn)
        self.activity_btn = ActivityButton()
        self.activity_btn.toggled.connect(self._set_activity_visible)
        header.addWidget(self.activity_btn)
        self.close_btn = CloseButton()
        self.close_btn.clicked.connect(self.dismiss)
        header.addWidget(self.close_btn)
        layout.addLayout(header)

        self.input = QtWidgets.QTextEdit()
        self.input.setPlaceholderText("Ask Axon...")
        self.input.setMinimumHeight(78)
        self.input.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.input.customContextMenuRequested.connect(self._input_menu)
        self.input.setStyleSheet(
            f"QTextEdit{{background:{SURFACE};color:{TEXT};border:1px solid {BORDER};"
            "border-radius:14px;padding:10px 48px 44px 12px;font-size:14px;"
            f"{FONT_CSS}selection-background-color:#5967f5;}}"
            "QTextEdit:focus{border:1px solid #4a4b55;background:#111217;}"
        )
        layout.addWidget(self.input, 1)

        self.send_btn = ArrowButton()
        self.send_btn.clicked.connect(self._on_send)
        self.send_btn.setParent(self.card)
        self.send_btn.raise_()

        self.mode_btn = ModeButton(self.card)  # sits bottom-left of the input, opposite Send
        self.mode_btn.raise_()
        self.approve_btn = ApprovalButton(self.card)  # second dropdown, next to mode
        self.approve_btn.raise_()
        # When a dropdown's label changes width (e.g. Agent -> Ask Maia), re-space the whole row.
        self.mode_btn.changed.connect(self._position_overlay_controls)
        self.approve_btn.changed.connect(self._position_overlay_controls)

        # Camera: attach a screenshot (whole screen / selected area) to ask Axon about it.
        self._attached_image = None
        self._attach_grown = False
        self.camera_btn = CameraButton(self.card)
        self.camera_btn.act_full.triggered.connect(self._capture_full)
        self.camera_btn.act_area.triggered.connect(self._capture_region)
        self.camera_btn.raise_()

        # Mic: click to start recording, click again to stop; then replay or remove the clip.
        self._recorder = None
        self._voice_path = None
        self.mic_btn = MicButton(self.card)
        self.mic_btn.setToolTip("Click to record, click again to stop")
        self.mic_btn.clicked.connect(self._toggle_record)
        self.mic_btn.raise_()
        self.voice_text.connect(self._on_voice_text)

        self.setAcceptDrops(True)  # drop a file onto the composer to act on it

        self.attach_bar = QtWidgets.QWidget(self.card)
        _ab = QtWidgets.QHBoxLayout(self.attach_bar)
        _ab.setContentsMargins(2, 0, 2, 0)
        _ab.setSpacing(8)
        self.attach_thumb = QtWidgets.QLabel()
        self.attach_thumb.setFixedSize(60, 38)
        self.attach_thumb.setScaledContents(True)
        self.attach_thumb.setStyleSheet("border:1px solid #2c2e38;border-radius:6px;background:#0d0e12;")
        _ab.addWidget(self.attach_thumb)
        self.attach_label = QtWidgets.QLabel("Screenshot attached — ask about it")
        self.attach_label.setStyleSheet(f"color:{MUTED};{FONT_CSS}font-size:12px;")
        _ab.addWidget(self.attach_label)
        _ab.addStretch(1)
        self.attach_remove = QtWidgets.QPushButton("✕")
        self.attach_remove.setCursor(QtCore.Qt.PointingHandCursor)
        self.attach_remove.setFixedSize(22, 22)
        self.attach_remove.setStyleSheet(
            f"QPushButton{{background:transparent;color:{MUTED};border:none;font-size:13px;}}"
            f"QPushButton:hover{{color:{TEXT};}}")
        self.attach_remove.clicked.connect(self.clear_attachment)
        _ab.addWidget(self.attach_remove)
        self.attach_bar.hide()
        layout.insertWidget(1, self.attach_bar)

        # Voice-clip bar (appears after recording): replay / remove the recording.
        self.voice_bar = QtWidgets.QWidget(self.card)
        _vb = QtWidgets.QHBoxLayout(self.voice_bar)
        _vb.setContentsMargins(2, 0, 2, 0)
        _vb.setSpacing(8)
        self.voice_play = QtWidgets.QPushButton("▶")  # ▶
        self.voice_play.setCursor(QtCore.Qt.PointingHandCursor)
        self.voice_play.setFixedSize(24, 24)
        self.voice_play.setStyleSheet(
            f"QPushButton{{background:transparent;color:{TEXT};border:1px solid {BORDER};border-radius:12px;}}"
            "QPushButton:hover{border-color:#4a4b55;}")
        self.voice_play.clicked.connect(self._play_voice)
        _vb.addWidget(self.voice_play)
        self.voice_label = QtWidgets.QLabel("Voice note")
        self.voice_label.setStyleSheet(f"color:{MUTED};{FONT_CSS}font-size:12px;")
        _vb.addWidget(self.voice_label)
        _vb.addStretch(1)
        self.voice_remove = QtWidgets.QPushButton("✕")  # ✕
        self.voice_remove.setCursor(QtCore.Qt.PointingHandCursor)
        self.voice_remove.setFixedSize(22, 22)
        self.voice_remove.setStyleSheet(
            f"QPushButton{{background:transparent;color:{MUTED};border:none;font-size:13px;}}"
            f"QPushButton:hover{{color:{TEXT};}}")
        self.voice_remove.clicked.connect(self._remove_voice)
        _vb.addWidget(self.voice_remove)
        self.voice_bar.hide()
        layout.insertWidget(2, self.voice_bar)

        self.log = QtWidgets.QTextEdit()
        self.log.setReadOnly(True)
        self.log.setStyleSheet(
            f"QTextEdit{{background:{SURFACE_2};color:#c8cad5;border:1px solid #292b35;"
            f"border-radius:12px;padding:10px 12px;{FONT_CSS}font-size:12px;line-height:1.35;}}"
        )
        self.log.hide()
        layout.addWidget(self.log, 1)
        self.size_grip = QtWidgets.QSizeGrip(self.card)
        self.size_grip.setFixedSize(16, 16)

        for drag_source in (self.card, mark, title):
            drag_source.installEventFilter(self)
        self.prepare_for_open()

    def prepare_for_open(self):
        self.ensurePolished()
        self.card.ensurePolished()
        if self.card.layout():
            self.card.layout().activate()
        self._refresh_min_height()
        self._position_overlay_controls()
        if not self._prepared_for_open:
            try:
                self.create()
            except Exception:
                pass
            self._prepared_for_open = True

    def _on_send(self):
        if self.send_btn.is_running():
            self.cancel_requested.emit()
            return

        text = self.input.toPlainText().strip()
        mode = self.mode_btn.mode()
        if mode == "guide":
            text = text or "Help me with what is on my screen right now."
            if not self._activity_visible:
                self.activity_btn.setChecked(True)
            self.log.clear()
            self.send_btn.set_running(True)
            self.guide_requested.emit(text)
        elif mode == "ask":
            if not text and not self._attached_image:
                return
            if not self._activity_visible:
                self.activity_btn.setChecked(True)
            self.log.clear()
            self.send_btn.set_running(True)
            self.ask_requested.emit(text)
        else:
            if not text and not self._attached_image:
                return
            if not self._activity_visible:
                self.activity_btn.setChecked(True)
            self.log.clear()
            self.send_btn.set_running(True)
            self.submitted.emit(text)

    # ---- screenshot attachment -------------------------------------------------
    def attached_image(self):
        return self._attached_image

    _ATTACH_EXTRA = 58  # extra height the attachment row needs above the input

    def clear_attachment(self, delete=True):
        if delete and self._attached_image:
            try:
                os.remove(self._attached_image)
            except Exception:
                pass
        self._attached_image = None
        if hasattr(self, "attach_bar"):
            self.attach_bar.hide()
        self._refresh_min_height()
        if not self._activity_visible and self.height() > self.minimumHeight():
            self.resize(self.width(), self.minimumHeight())  # give back the borrowed row
        self._position_overlay_controls()

    def _attach_pixmap(self, pix):
        import tempfile
        self.clear_attachment(delete=True)  # drop any previous shot
        fd, path = tempfile.mkstemp(prefix="axon_shot_", suffix=".png")
        os.close(fd)
        pix.save(path, "PNG")
        self._attached_image = path
        self.attach_thumb.setPixmap(
            pix.scaled(self.attach_thumb.size() * 2, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation))
        self.attach_bar.show()
        self._refresh_min_height()                       # raising the minimum auto-grows the panel
        if self.height() < self.minimumHeight():
            self.resize(self.width(), self.minimumHeight())
        self._position_overlay_controls()

    def _capture_full(self):
        scr = self.screen()
        self.hide()  # keep the composer out of the shot
        QtCore.QTimer.singleShot(220, lambda: self._do_full(scr))

    def _do_full(self, scr):
        try:
            self._attach_pixmap(scr.grabWindow(0))
        except Exception:
            pass
        self.show()
        self.raise_()
        self.activateWindow()

    def _capture_region(self):
        scr = self.screen()
        self.hide()
        QtCore.QTimer.singleShot(220, lambda: self._show_selector(scr))

    def _show_selector(self, scr):
        try:
            full = scr.grabWindow(0)
        except Exception:
            self.show()
            return
        self._selector = RegionSelector(scr, full)
        self._selector.setFocusPolicy(QtCore.Qt.StrongFocus)
        self._selector.selected.connect(self._on_region)
        self._selector.show()
        self._selector.raise_()
        self._selector.activateWindow()
        self._selector.setFocus()

    def _on_region(self, pix):
        if pix is not None and not pix.isNull():
            self._attach_pixmap(pix)
        self.show()
        self.raise_()
        self.activateWindow()

    def _refresh_min_height(self):
        """Minimum height that always fits the header, input and the bottom control row, plus any
        optional rows currently shown (attachment thumbnail, activity log). Raising the minimum
        auto-grows the panel; we shrink to it when rows go away — so resizing never clips controls."""
        h = 200  # header + input + control row + margins
        if self._attached_image is not None:
            h += self._ATTACH_EXTRA
        if self._voice_path is not None:
            h += 40  # voice-clip bar
        if self._activity_visible:
            h += 130
        self.setMinimumHeight(h)

    def _set_activity_visible(self, visible):
        self._activity_visible = visible
        self.log.setVisible(visible)
        self._refresh_min_height()
        if visible:
            if self.height() < self.minimumHeight():
                self.resize(self.width(), self.minimumHeight())
        elif self.height() > self.minimumHeight():
            self.resize(self.width(), self.minimumHeight())

    def showEvent(self, event):
        super().showEvent(event)
        self.prepare_for_open()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if not hasattr(self, "card"):
            return
        self.card.setGeometry(8, 8, self.width() - 16, self.height() - 16)
        self._position_overlay_controls()

    def _position_overlay_controls(self):
        if not hasattr(self, "card"):
            return
        if hasattr(self, "size_grip"):
            self.size_grip.move(
                self.card.width() - self.size_grip.width() - 6,
                self.card.height() - self.size_grip.height() - 6,
            )
        if hasattr(self, "send_btn") and hasattr(self, "input"):
            input_geometry = self.input.geometry()
            self.send_btn.move(
                input_geometry.right() - self.send_btn.width() - 12,
                input_geometry.bottom() - self.send_btn.height() - 10,
            )
            self.send_btn.raise_()
            if hasattr(self, "mode_btn"):
                row_y = self.send_btn.y() + (self.send_btn.height() - self.mode_btn.height()) // 2
                self.mode_btn.move(input_geometry.left() + 12, row_y)
                self.mode_btn.raise_()
                if hasattr(self, "approve_btn"):
                    self.approve_btn.move(
                        self.mode_btn.x() + self.mode_btn.width() + 8, row_y)
                    self.approve_btn.raise_()
                    if hasattr(self, "camera_btn"):
                        cam_y = self.send_btn.y() + (self.send_btn.height() - self.camera_btn.height()) // 2
                        self.camera_btn.move(
                            self.approve_btn.x() + self.approve_btn.width() + 6, cam_y)
                        self.camera_btn.raise_()
                        if hasattr(self, "mic_btn"):
                            self.mic_btn.move(
                                self.camera_btn.x() + self.camera_btn.width() + 4, cam_y)
                            self.mic_btn.raise_()

    def _toggle_record(self):
        """Click to start recording; click again to stop, then transcribe + show the clip."""
        from axon.vision import Recorder
        if self._recorder is None:
            rec = Recorder()
            if rec.start():
                self._recorder = rec
                self.mic_btn.set_recording(True)
            return
        rec, self._recorder = self._recorder, None
        self.mic_btn.set_recording(False)
        path = rec.stop()
        if not path:
            return
        self._voice_path = path
        self.voice_bar.show()
        self._refresh_min_height()
        from axon.vision import is_silent
        if is_silent(path):
            self.voice_label.setText("No mic audio — check your microphone is on")
            return
        self.voice_label.setText("Voice note — transcribing…")
        import threading

        def work():
            from axon.vision import transcribe_audio
            self.voice_text.emit(transcribe_audio(path) or "")

        threading.Thread(target=work, daemon=True).start()

    def _on_voice_text(self, txt):
        if txt:
            cur = self.input.toPlainText().strip()
            self.input.setPlainText((cur + " " + txt).strip() if cur else txt)
            self.voice_label.setText("Voice note")
        else:
            self.voice_label.setText("Voice note (couldn't transcribe)")
        self.input.setFocus()

    def _play_voice(self):
        from axon.vision import play_audio
        if self._voice_path:
            play_audio(self._voice_path)

    def _remove_voice(self):
        from axon.vision import stop_audio
        stop_audio()
        if self._voice_path:
            try:
                os.remove(self._voice_path)
            except Exception:
                pass
        self._voice_path = None
        self.voice_bar.hide()
        self._refresh_min_height()

    def _input_menu(self, pos):
        """Right-click menu on the input: the normal copy/paste plus a Quick actions submenu."""
        menu = self.input.createStandardContextMenu()
        menu.addSeparator()
        qa = menu.addMenu("Quick actions")
        for label, text in self.QUICK_ACTIONS:
            act = qa.addAction(label)
            act.triggered.connect(
                lambda _=False, t=text: (self.input.setPlainText(t),
                                         self.input.moveCursor(QtGui.QTextCursor.End),
                                         self.input.setFocus()))
        menu.exec(self.input.mapToGlobal(pos))

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        paths = [u.toLocalFile() for u in e.mimeData().urls() if u.toLocalFile()]
        if not paths:
            return
        cur = self.input.toPlainText().strip()
        ref = '"' + paths[0] + '"'
        self.input.setPlainText((cur + " " + ref).strip() if cur else (ref + " — "))
        self.input.setFocus()
        e.acceptProposedAction()

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton and event.position().y() <= 62:
            self._drag_offset = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._drag_offset is not None:
            self.move(event.globalPosition().toPoint() - self._drag_offset)
            event.accept()
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self._drag_offset = None
        super().mouseReleaseEvent(event)

    def dismiss(self):
        self.hide()
        self.dismissed.emit()

    def eventFilter(self, obj, event):
        event_type = event.type()
        if event_type == QtCore.QEvent.MouseButtonPress and event.button() == QtCore.Qt.LeftButton:
            if obj is self.card and event.position().y() > 62:
                return False
            self._drag_offset = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            return True
        if event_type == QtCore.QEvent.MouseMove and self._drag_offset is not None:
            self.move(event.globalPosition().toPoint() - self._drag_offset)
            return True
        if event_type == QtCore.QEvent.MouseButtonRelease and event.button() == QtCore.Qt.LeftButton:
            self._drag_offset = None
            return True
        return super().eventFilter(obj, event)

    def append_log(self, line):
        self.log.append(_md_to_html(line))
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())

    def task_finished(self):
        self.send_btn.set_running(False)

    def task_started(self):
        self.send_btn.set_running(True)

    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Escape:
            self.dismiss()
        elif event.key() in (QtCore.Qt.Key_Return, QtCore.Qt.Key_Enter) and (
            event.modifiers() & QtCore.Qt.ControlModifier
        ):
            self._on_send()
        else:
            super().keyPressEvent(event)
