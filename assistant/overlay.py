"""
Floating always-on-top dot + composer for the Axon intelligence computer-use agent.

- A small draggable circular dot floats on top of everything.
- Click it without dragging to open a composer panel.
- Type a request, press Send or Ctrl+Enter; the agent runs in a background
  thread and live status streams into the panel.

Run:  python overlay.py
"""

import socket
import os
import sys
import time
import math
import threading

from PySide6 import QtCore, QtGui, QtWidgets

import agent

# Single-instance guard: hold a localhost port for the app's lifetime. If another
# Axon intelligence dot is already running, the bind fails and this instance exits.
_INSTANCE_LOCK = None

ACCENT = "#ffffff"
ACCENT_2 = "#cfd2dc"
PANEL_BG = "#050506"
PANEL_BG_2 = "#0b0b0e"
SURFACE = "#101014"
SURFACE_2 = "#08090c"
BORDER = "#282932"
TEXT = "#f7f7fb"
MUTED = "#9da0ad"
FONT_FAMILY = "Arial"
FONT_CSS = f"font-family:{FONT_FAMILY};"
HEADER_ICON_SIZE = 22
CONTROL_ICON_SIZE = 30


def _acquire_single_instance(port=49737):
    global _INSTANCE_LOCK
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", port))
        s.listen(1)
    except OSError:
        return False
    _INSTANCE_LOCK = s
    return True


class AgentWorker(QtCore.QObject):
    status = QtCore.Signal(str)
    finished = QtCore.Signal(str)
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
                )
            self.finished.emit(result or "Done.")
        except Exception as e:
            self.finished.emit(f"Error: {e}")


class ApprovalPopup(QtWidgets.QWidget):
    """Small always-on-top prompt asking the user to Approve or Skip a pending action."""
    decided = QtCore.Signal(bool)

    def __init__(self):
        super().__init__()
        self.setWindowFlags(
            QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.Tool
        )
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.resize(450, 172)
        self.card = PanelFrame(self)
        lay = QtWidgets.QVBoxLayout(self.card)
        lay.setContentsMargins(20, 18, 20, 16)
        lay.setSpacing(10)
        title = QtWidgets.QLabel("Approve this action?")
        title.setStyleSheet(f"color:{TEXT};{FONT_CSS}font-size:15px;font-weight:700;")
        lay.addWidget(title)
        self.msg = QtWidgets.QLabel("")
        self.msg.setWordWrap(True)
        self.msg.setStyleSheet(f"color:{ACCENT_2};{FONT_CSS}font-size:13px;")
        lay.addWidget(self.msg, 1)
        row = QtWidgets.QHBoxLayout()
        row.addStretch(1)
        skip = QtWidgets.QPushButton("Skip")
        approve = QtWidgets.QPushButton("Approve")
        for b in (skip, approve):
            b.setCursor(QtCore.Qt.PointingHandCursor)
        skip.setStyleSheet(
            f"QPushButton{{background:{SURFACE};color:{TEXT};border:1px solid {BORDER};"
            f"border-radius:10px;padding:8px 18px;{FONT_CSS}font-size:13px;font-weight:600;}}"
            "QPushButton:hover{background:#1a1b22;}")
        approve.setStyleSheet(
            f"QPushButton{{background:#ffffff;color:#0a0a0c;border:none;"
            f"border-radius:10px;padding:8px 20px;{FONT_CSS}font-size:13px;font-weight:700;}}"
            "QPushButton:hover{background:#e9e9ee;}")
        skip.clicked.connect(lambda: self._decide(False))
        approve.clicked.connect(lambda: self._decide(True))
        row.addWidget(skip)
        row.addWidget(approve)
        lay.addLayout(row)
        self._reposition()

    def _reposition(self):
        self.card.setGeometry(8, 8, self.width() - 16, self.height() - 16)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._reposition()

    def ask(self, description, screen=None):
        self.msg.setText(description)
        geo = (screen.availableGeometry() if screen is not None
               else QtWidgets.QApplication.primaryScreen().availableGeometry())
        self.move(geo.center().x() - self.width() // 2, geo.top() + 120)
        self.show()
        self.raise_()
        self.activateWindow()

    def _decide(self, approved):
        self.hide()
        self.decided.emit(approved)


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


class HighlightOverlay(QtWidgets.QWidget):
    """Full-desktop, click-through overlay that points at a spot the guide suggests."""

    def __init__(self):
        super().__init__()
        self.setWindowFlags(
            QtCore.Qt.FramelessWindowHint
            | QtCore.Qt.WindowStaysOnTopHint
            | QtCore.Qt.Tool
            | QtCore.Qt.WindowTransparentForInput
        )
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setAttribute(QtCore.Qt.WA_TransparentForMouseEvents)
        self._pt = None
        self._box = None
        self._click = None
        self._phase = 0.0
        self._timer = QtCore.QTimer(self)
        self._timer.setInterval(33)
        self._timer.timeout.connect(self._tick)
        self._hide_timer = QtCore.QTimer(self)
        self._hide_timer.setSingleShot(True)
        self._hide_timer.timeout.connect(self.hide)

    def _geo(self, screen):
        # Fractions are relative to ONE monitor (the dot's). Draw on that screen's geometry so the
        # mapping is exact regardless of multi-monitor layout / DPI.
        if screen is not None:
            return screen.geometry()
        return QtWidgets.QApplication.primaryScreen().virtualGeometry()

    def show_pointer(self, fx, fy, label, persist=False, hide_after_ms=12000, screen=None):
        vg = self._geo(screen)
        self.setGeometry(vg)
        fx = min(max(float(fx), 0.0), 1.0)
        fy = min(max(float(fy), 0.0), 1.0)
        self._box = None
        self._click = None
        self._pt = (fx * vg.width(), fy * vg.height(), label or "")
        self._phase = 0.0
        self.show()
        self.raise_()
        self._timer.start()
        self._hide_timer.stop()
        if not persist:
            self._hide_timer.start(hide_after_ms)

    def show_box(self, fx, fy, fw, fh, label, persist=False, hide_after_ms=12000, screen=None):
        vg = self._geo(screen)
        self.setGeometry(vg)
        x = max(0.0, min(float(fx), 1.0)) * vg.width()
        y = max(0.0, min(float(fy), 1.0)) * vg.height()
        w = max(0.0, float(fw)) * vg.width()
        h = max(0.0, float(fh)) * vg.height()
        self._pt = None
        self._click = None
        self._box = (x, y, w, h, label or "")
        self._phase = 0.0
        self.show()
        self.raise_()
        self._timer.stop()  # brackets are static — no pulsing/jitter
        self.update()
        self._hide_timer.stop()
        if not persist:
            self._hide_timer.start(hide_after_ms)

    def show_click(self, fx, fy, label, persist=False, hide_after_ms=12000, screen=None):
        vg = self._geo(screen)
        self.setGeometry(vg)
        fx = min(max(float(fx), 0.0), 1.0)
        fy = min(max(float(fy), 0.0), 1.0)
        self._pt = None
        self._box = None
        self._click = (fx * vg.width(), fy * vg.height(), label or "Click here")
        self._phase = 0.0
        self.show()
        self.raise_()
        self._timer.start()  # gentle pulse to draw the eye to the callout
        self._hide_timer.stop()
        if not persist:
            self._hide_timer.start(hide_after_ms)

    def hide_after(self, ms):
        self._hide_timer.start(ms)

    def hide(self):
        self._timer.stop()
        self._pt = None
        self._box = None
        self._click = None
        super().hide()

    def _tick(self):
        self._phase = (self._phase + 0.06) % (2 * math.pi)
        self.update()

    def _draw_label(self, p, label, bx, by, anchor_right_limit=None):
        f = QtGui.QFont(FONT_FAMILY)
        f.setPointSize(10)
        f.setBold(True)
        p.setFont(f)
        fm = p.fontMetrics()
        bw = fm.horizontalAdvance(label) + 18
        bh = fm.height() + 10
        if anchor_right_limit is not None and bx + bw > anchor_right_limit:
            bx = anchor_right_limit - bw
        if bx < 0:
            bx = 0
        if by < 0:
            by = 0
        cap = QtCore.QRectF(bx, by, bw, bh)
        # Solid black chip with a white outline -> readable on any background.
        p.setBrush(QtGui.QColor(0, 0, 0, 240))
        p.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 235), 1.5))
        p.drawRoundedRect(cap, 8, 8)
        p.setPen(QtGui.QColor(255, 255, 255))
        p.drawText(cap, QtCore.Qt.AlignCenter, label)

    def _bw_stroke(self, p, draw, pulse):
        """Draw a shape twice — a white halo then a black core — so it shows on any background."""
        p.setBrush(QtCore.Qt.NoBrush)
        p.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 240), 6))
        draw()
        p.setPen(QtGui.QPen(QtGui.QColor(0, 0, 0, 250), 3))
        draw()

    def paintEvent(self, _):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        pulse = (1 + math.sin(self._phase)) / 2
        if self._click:
            x, y, label = self._click
            # small target dot at the spot
            r = 6 + 3 * pulse
            self._bw_stroke(p, lambda: p.drawEllipse(QtCore.QPointF(x, y), r, r), pulse)
            # a "Click here" chip above-right, with a connector down to the dot
            f = QtGui.QFont(FONT_FAMILY)
            f.setPointSize(11)
            f.setBold(True)
            p.setFont(f)
            fm = p.fontMetrics()
            bw = fm.horizontalAdvance(label) + 22
            bh = fm.height() + 12
            bx = x + 22
            by = y - bh - 22
            if bx + bw > self.width():
                bx = x - bw - 22
            if by < 0:
                by = y + 22
            self._bw_stroke(p, lambda: p.drawLine(QtCore.QPointF(x, y),
                                                  QtCore.QPointF(bx + 12, by + bh / 2)), pulse)
            cap = QtCore.QRectF(bx, by, bw, bh)
            p.setBrush(QtGui.QColor(0, 0, 0, 240))
            p.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 240), 1.5))
            p.drawRoundedRect(cap, 9, 9)
            p.setPen(QtGui.QColor("#ffffff"))
            p.drawText(cap, QtCore.Qt.AlignCenter, label)
            return
        if self._box:
            x, y, w, h, label = self._box
            w = max(w, 10)
            h = max(h, 8)
            # Fixed breathing room (NO pulse) so the brackets stay perfectly still on the target.
            gap = 3
            x -= gap
            y -= gap
            w += 2 * gap
            h += 2 * gap
            arm = max(min(h * 0.3, 14), 5)  # bracket "feet" scale down for short/small buttons

            def poly(pts):
                return QtGui.QPolygonF([QtCore.QPointF(px, py) for px, py in pts])

            left = poly([(x + arm, y), (x, y), (x, y + h), (x + arm, y + h)])    # [
            right = poly([(x + w - arm, y), (x + w, y), (x + w, y + h), (x + w - arm, y + h)])  # ]

            def draw():
                p.drawPolyline(left)
                p.drawPolyline(right)

            self._bw_stroke(p, draw, pulse)
            if label:
                self._draw_label(p, label, x, y - 30, anchor_right_limit=self.width())
            return
        if self._pt:
            x, y, label = self._pt
            r = 26 + 8 * pulse
            c = QtCore.QPointF(x, y)
            self._bw_stroke(p, lambda: p.drawEllipse(c, r, r), pulse)
            self._bw_stroke(p, lambda: p.drawEllipse(c, 15, 15), pulse)
            if label:
                self._draw_label(p, label, x + r + 8, y - 17, anchor_right_limit=self.width())


class PanelFrame(QtWidgets.QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(QtCore.Qt.WA_StyledBackground, False)

    def paintEvent(self, _):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        rect = QtCore.QRectF(0.5, 0.5, self.width() - 1, self.height() - 1)
        path = QtGui.QPainterPath()
        path.addRoundedRect(rect, 18, 18)

        p.fillPath(path, QtGui.QColor(PANEL_BG))

        highlight = QtGui.QLinearGradient(rect.topLeft(), rect.bottomLeft())
        highlight.setColorAt(0, QtGui.QColor(255, 255, 255, 30))
        highlight.setColorAt(1, QtGui.QColor(255, 255, 255, 6))
        p.setPen(QtGui.QPen(QtGui.QBrush(highlight), 1))
        p.drawPath(path)


class MarkView(QtWidgets.QWidget):
    def __init__(self, size=HEADER_ICON_SIZE, parent=None):
        super().__init__(parent)
        self.setFixedSize(size, size)

    def paintEvent(self, _):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        cx, cy = self.width() / 2, self.height() / 2
        radius = min(self.width(), self.height()) * 0.38
        c1 = radius * 0.21
        c2 = radius * 0.40
        star = QtGui.QPainterPath()
        star.moveTo(cx, cy - radius)
        star.cubicTo(cx + c1, cy - c2, cx + c2, cy - c1, cx + radius, cy)
        star.cubicTo(cx + c2, cy + c1, cx + c1, cy + c2, cx, cy + radius)
        star.cubicTo(cx - c1, cy + c2, cx - c2, cy + c1, cx - radius, cy)
        star.cubicTo(cx - c2, cy - c1, cx - c1, cy - c2, cx, cy - radius)
        p.setBrush(QtGui.QColor("#ffffff"))
        p.setPen(QtCore.Qt.NoPen)
        p.drawPath(star)


class IconButton(QtWidgets.QPushButton):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(CONTROL_ICON_SIZE, CONTROL_ICON_SIZE)
        self.setCursor(QtCore.Qt.PointingHandCursor)

    def enterEvent(self, event):
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.update()
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        super().mousePressEvent(event)
        self.update()

    def mouseReleaseEvent(self, event):
        super().mouseReleaseEvent(event)
        self.update()


class ArrowButton(IconButton):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._running = False
        self.setAccessibleName("Run")

    def is_running(self):
        return self._running

    def set_running(self, running):
        if self._running == running:
            return
        self._running = running
        self.setAccessibleName("Stop" if running else "Run")
        self.update()

    def paintEvent(self, _):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        inset = 1.0 if not self.isDown() else 2.0
        rect = QtCore.QRectF(inset, inset, self.width() - (inset * 2), self.height() - (inset * 2))
        bg = QtGui.QColor("#ffffff" if self.isEnabled() else "#34343a")
        if self.isDown() and self.isEnabled():
            bg = QtGui.QColor("#e8e9ee")
        elif self.underMouse() and self.isEnabled():
            bg = QtGui.QColor("#f6f7fb")
        p.setBrush(bg)
        p.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 42), 1))
        p.drawEllipse(rect)

        offset = 1.0 if self.isDown() else 0.0
        cx, cy = self.width() / 2, (self.height() / 2) + offset
        glyph = QtGui.QColor("#050506" if self.isEnabled() else "#8d8f99")
        if self._running:
            side = self.width() * 0.28
            stop_rect = QtCore.QRectF(cx - side / 2, cy - side / 2, side, side)
            p.setBrush(glyph)
            p.setPen(QtCore.Qt.NoPen)
            p.drawRoundedRect(stop_rect, 2.0, 2.0)
        else:
            pen = QtGui.QPen(glyph, 1.9)
            pen.setCapStyle(QtCore.Qt.RoundCap)
            pen.setJoinStyle(QtCore.Qt.RoundJoin)
            p.setPen(pen)
            shaft = self.height() * 0.36
            head = self.height() * 0.16
            p.drawLine(QtCore.QPointF(cx, cy + shaft / 2), QtCore.QPointF(cx, cy - shaft / 2))
            p.drawLine(QtCore.QPointF(cx, cy - shaft / 2), QtCore.QPointF(cx - head, cy - shaft / 2 + head))
            p.drawLine(QtCore.QPointF(cx, cy - shaft / 2), QtCore.QPointF(cx + head, cy - shaft / 2 + head))


class ActivityButton(IconButton):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCheckable(True)
        self.setAccessibleName("Toggle activity")

    def paintEvent(self, _):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        inset = 1.0 if not self.isDown() else 2.0
        rect = QtCore.QRectF(inset, inset, self.width() - (inset * 2), self.height() - (inset * 2))
        if self.isDown():
            bg = QtGui.QColor("#22242b")
        elif self.isChecked() or self.underMouse():
            bg = QtGui.QColor("#18191f")
        else:
            bg = QtGui.QColor("#0d0e12")
        p.setBrush(bg)
        p.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 40), 1))
        p.drawRoundedRect(rect, self.width() * 0.32, self.height() * 0.32)

        pen = QtGui.QPen(QtGui.QColor("#f1f2f6" if self.isChecked() else "#d9dbe3"), 1.6)
        pen.setCapStyle(QtCore.Qt.RoundCap)
        p.setPen(pen)
        left = self.width() * 0.34
        right = self.width() * 0.66
        for y in (0.40, 0.50, 0.60):
            y_pos = self.height() * y
            p.drawLine(QtCore.QPointF(left, y_pos), QtCore.QPointF(right, y_pos))


class CloseButton(IconButton):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAccessibleName("Close composer")

    def paintEvent(self, _):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        inset = 1.0 if not self.isDown() else 2.0
        rect = QtCore.QRectF(inset, inset, self.width() - (inset * 2), self.height() - (inset * 2))
        bg = QtGui.QColor("#24252c" if self.isDown() else ("#1c1d23" if self.underMouse() else "#0d0e12"))
        p.setBrush(bg)
        p.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 42), 1))
        p.drawRoundedRect(rect, self.width() * 0.32, self.height() * 0.32)

        pen = QtGui.QPen(QtGui.QColor("#ffffff" if self.underMouse() else "#e8e9ee"), 1.8)
        pen.setCapStyle(QtCore.Qt.RoundCap)
        p.setPen(pen)
        pad = self.width() * 0.37
        p.drawLine(QtCore.QPointF(pad, pad), QtCore.QPointF(self.width() - pad, self.height() - pad))
        p.drawLine(QtCore.QPointF(self.width() - pad, pad), QtCore.QPointF(pad, self.height() - pad))


class GuideButton(IconButton):
    """Coach-me mode: a '?' button that asks the guide to look at the screen and advise."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAccessibleName("Guide me")

    def paintEvent(self, _):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        inset = 1.0 if not self.isDown() else 2.0
        rect = QtCore.QRectF(inset, inset, self.width() - (inset * 2), self.height() - (inset * 2))
        if self.isDown():
            bg = QtGui.QColor("#22242b")
        elif self.underMouse():
            bg = QtGui.QColor("#18191f")
        else:
            bg = QtGui.QColor("#0d0e12")
        p.setBrush(bg)
        p.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 40), 1))
        p.drawRoundedRect(rect, self.width() * 0.32, self.height() * 0.32)
        p.setPen(QtGui.QColor("#f1f2f6" if self.underMouse() else "#d9dbe3"))
        f = QtGui.QFont(FONT_FAMILY)
        f.setPointSize(max(8, int(self.height() * 0.40)))
        f.setBold(True)
        p.setFont(f)
        p.drawText(QtCore.QRectF(0, 0, self.width(), self.height()), QtCore.Qt.AlignCenter, "?")


class ModeButton(QtWidgets.QPushButton):
    """Dropdown to choose how Axon responds: Agent (does it) or Guide me (coaches you)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self.setFixedHeight(26)
        self._mode = "agent"
        menu = QtWidgets.QMenu(self)
        menu.setStyleSheet(
            f"QMenu{{background:{SURFACE};color:{TEXT};border:1px solid {BORDER};"
            f"border-radius:10px;padding:6px;{FONT_CSS}font-size:12px;}}"
            "QMenu::item{padding:7px 12px;border-radius:7px;}"
            "QMenu::item:selected{background:#23252e;}"
        )
        a_agent = menu.addAction("Agent  —  do it for me")
        a_guide = menu.addAction("Guide me  —  coach me")
        a_ask = menu.addAction("Ask Maia  —  just answer")
        a_agent.triggered.connect(lambda: self.set_mode("agent"))
        a_guide.triggered.connect(lambda: self.set_mode("guide"))
        a_ask.triggered.connect(lambda: self.set_mode("ask"))
        self.setMenu(menu)
        self.setStyleSheet(
            f"QPushButton{{background:transparent;color:{MUTED};border:none;"
            f"padding:3px 4px;{FONT_CSS}font-size:12px;font-weight:600;}}"
            f"QPushButton:hover{{color:{TEXT};}}"
            "QPushButton::menu-indicator{width:0px;}"
        )
        self._refresh()

    def mode(self):
        return self._mode

    def set_mode(self, m):
        self._mode = m
        self._refresh()

    def _refresh(self):
        label = {"agent": "Agent", "guide": "Guide me", "ask": "Ask Maia"}.get(self._mode, "Agent")
        self.setText(f"{label}  ▾")
        self.adjustSize()


class ApprovalButton(QtWidgets.QPushButton):
    """Separate dropdown: Ask before acting, or act automatically."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self.setFixedHeight(26)
        self._ask = True
        menu = QtWidgets.QMenu(self)
        menu.setStyleSheet(
            f"QMenu{{background:{SURFACE};color:{TEXT};border:1px solid {BORDER};"
            f"border-radius:10px;padding:6px;{FONT_CSS}font-size:12px;}}"
            "QMenu::item{padding:7px 12px;border-radius:7px;}"
            "QMenu::item:selected{background:#23252e;}"
        )
        a_ask = menu.addAction("Ask before acting")
        a_auto = menu.addAction("Act automatically")
        a_ask.triggered.connect(lambda: self.set_ask(True))
        a_auto.triggered.connect(lambda: self.set_ask(False))
        self.setMenu(menu)
        self.setStyleSheet(
            f"QPushButton{{background:transparent;color:{MUTED};border:none;"
            f"padding:3px 4px;{FONT_CSS}font-size:12px;font-weight:600;}}"
            f"QPushButton:hover{{color:{TEXT};}}"
            "QPushButton::menu-indicator{width:0px;}"
        )
        self._refresh()

    def ask_before(self):
        return self._ask

    def set_ask(self, v):
        self._ask = bool(v)
        self._refresh()

    def _refresh(self):
        self.setText(("Ask" if self._ask else "Auto") + "  ▾")
        self.adjustSize()


class RegionSelector(QtWidgets.QWidget):
    """Snipping-tool overlay: shows a frozen screenshot, the user drags a rectangle, and the
    selected region is emitted as a QPixmap (or None if cancelled)."""
    selected = QtCore.Signal(object)  # QPixmap of the chosen area, or None

    def __init__(self, screen, pixmap):
        super().__init__()
        self._pix = pixmap
        self._dpr = pixmap.devicePixelRatio() or 1.0
        self.setWindowFlags(
            QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.Tool
        )
        self.setGeometry(screen.geometry())
        self.setCursor(QtCore.Qt.CrossCursor)
        self._origin = None
        self._cur = None

    def paintEvent(self, e):
        p = QtGui.QPainter(self)
        p.drawPixmap(self.rect(), self._pix)              # the frozen screenshot (bright)
        dim = QtGui.QColor(0, 0, 0, 120)
        full = self.rect()
        if self._origin and self._cur:
            r = QtCore.QRect(self._origin, self._cur).normalized()
            # Dim only the four areas AROUND the selection — the selection itself stays bright.
            p.fillRect(QtCore.QRect(0, 0, full.width(), r.top()), dim)                       # above
            p.fillRect(QtCore.QRect(0, r.bottom(), full.width(), full.height() - r.bottom()), dim)  # below
            p.fillRect(QtCore.QRect(0, r.top(), r.left(), r.height()), dim)                  # left
            p.fillRect(QtCore.QRect(r.right(), r.top(), full.width() - r.right(), r.height()), dim)  # right
            p.setPen(QtGui.QPen(QtGui.QColor("#5967f5"), 2))
            p.drawRect(r)
        else:
            p.fillRect(full, dim)
            p.setPen(QtGui.QColor("#e8e8ee"))
            f = p.font(); f.setPointSize(12); p.setFont(f)
            p.drawText(full, QtCore.Qt.AlignHCenter | QtCore.Qt.AlignTop,
                       "\n\nDrag to select an area  ·  Esc to cancel")
        p.end()

    def mousePressEvent(self, e):
        if e.button() == QtCore.Qt.LeftButton:
            self._origin = e.position().toPoint()
            self._cur = self._origin
            self.update()

    def mouseMoveEvent(self, e):
        if self._origin is not None:
            self._cur = e.position().toPoint()
            self.update()

    def mouseReleaseEvent(self, e):
        if e.button() != QtCore.Qt.LeftButton or self._origin is None:
            return
        r = QtCore.QRect(self._origin, e.position().toPoint()).normalized()
        self.hide()
        if r.width() > 4 and r.height() > 4:
            dev = QtCore.QRect(int(r.x() * self._dpr), int(r.y() * self._dpr),
                               int(r.width() * self._dpr), int(r.height() * self._dpr))
            self.selected.emit(self._pix.copy(dev))
        else:
            self.selected.emit(None)
        self.close()

    def keyPressEvent(self, e):
        if e.key() == QtCore.Qt.Key_Escape:
            self.hide()
            self.selected.emit(None)
            self.close()


class CameraButton(QtWidgets.QPushButton):
    """Camera icon next to Send: capture the whole screen or a selected area to attach + ask about."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self.setFixedSize(30, 26)
        self.setToolTip("Attach a screenshot to ask about")
        menu = QtWidgets.QMenu(self)
        menu.setStyleSheet(
            f"QMenu{{background:{SURFACE};color:{TEXT};border:1px solid {BORDER};"
            f"border-radius:10px;padding:6px;{FONT_CSS}font-size:12px;}}"
            "QMenu::item{padding:7px 12px;border-radius:7px;}"
            "QMenu::item:selected{background:#23252e;}"
        )
        self.act_full = menu.addAction("Entire screen")
        self.act_area = menu.addAction("Select area")
        self.setMenu(menu)
        self.setStyleSheet(
            "QPushButton{background:transparent;border:none;}"
            "QPushButton::menu-indicator{width:0px;}"
        )

    def paintEvent(self, e):
        super().paintEvent(e)
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        hovered = self.underMouse()
        col = QtGui.QColor(TEXT if hovered else MUTED)
        p.setPen(QtGui.QPen(col, 1.6))
        w, h = self.width(), self.height()
        body = QtCore.QRectF(w / 2 - 8, h / 2 - 5, 16, 11)
        p.drawRoundedRect(body, 2.5, 2.5)
        # little viewfinder bump on top
        p.drawLine(int(w / 2 - 3), int(h / 2 - 5), int(w / 2 - 1), int(h / 2 - 7))
        p.drawLine(int(w / 2 - 1), int(h / 2 - 7), int(w / 2 + 2), int(h / 2 - 7))
        p.setBrush(QtCore.Qt.NoBrush)
        p.drawEllipse(QtCore.QPointF(w / 2, h / 2 + 0.5), 3.0, 3.0)
        p.end()


class Composer(QtWidgets.QWidget):
    submitted = QtCore.Signal(str)
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
        self.setMinimumSize(320, 176)
        self.resize(420, 208)
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

        # Camera: attach a screenshot (whole screen / selected area) to ask Axon about it.
        self._attached_image = None
        self.camera_btn = CameraButton(self.card)
        self.camera_btn.act_full.triggered.connect(self._capture_full)
        self.camera_btn.act_area.triggered.connect(self._capture_region)
        self.camera_btn.raise_()

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

        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setStyleSheet(
            f"QPlainTextEdit{{background:{SURFACE_2};color:#c8cad5;border:1px solid #292b35;"
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

    def clear_attachment(self, delete=True):
        if delete and self._attached_image:
            try:
                os.remove(self._attached_image)
            except Exception:
                pass
        self._attached_image = None
        if hasattr(self, "attach_bar"):
            self.attach_bar.hide()

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

    def _set_activity_visible(self, visible):
        self._activity_visible = visible
        if visible:
            self._height_before_activity = self.height()
            self.log.setVisible(True)
            if self.height() < 330:
                self.resize(self.width(), 330)
        else:
            self.log.setVisible(False)
            target_height = max(self.minimumHeight(), self._height_before_activity)
            if self.height() > target_height:
                self.resize(self.width(), target_height)

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
        self.log.appendPlainText(line)
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


class FloatingDot(QtWidgets.QWidget):
    DIAM = 44

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

        # Tell the backend which window is the dot, so opened apps land on the dot's monitor.
        try:
            agent.DOT_HWND = int(self.winId())
            agent.DOT_PID = os.getpid()
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
            agent.DOT_HWND = int(ref.winId())
            agent.DOT_PID = os.getpid()
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

    @QtCore.Slot(str)
    def _on_finished(self, result):
        self._approval_popup.hide()
        self.composer.append_log("")
        self.composer.append_log(result)
        self.composer.task_finished()
        w = self._worker
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


def main():
    if not _acquire_single_instance():
        print("Axon intelligence is already running - not starting a second dot.")
        return
    app = QtWidgets.QApplication(sys.argv)
    app.setFont(QtGui.QFont(FONT_FAMILY, 9))
    app.setQuitOnLastWindowClosed(False)
    dot = FloatingDot()
    dot.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
