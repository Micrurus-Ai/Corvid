"""Reusable buttons and the mode / approval / camera controls."""
from PySide6 import QtCore, QtGui, QtWidgets

import agent
from axon.settings import load_settings, save_settings
from axon.ui.theme import (ACCENT, ACCENT_2, PANEL_BG, PANEL_BG_2, SURFACE, BORDER, TEXT,
                            MUTED, FONT_FAMILY, FONT_CSS, HEADER_ICON_SIZE, CONTROL_ICON_SIZE)


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


class ProfileButton(IconButton):
    """Profile / settings entry (sits before the activity icon). Opens the profile panel where the
    inbox auto-filer and other preferences are toggled."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAccessibleName("Profile and settings")
        menu = QtWidgets.QMenu(self)
        menu.setStyleSheet(
            f"QMenu{{background:{SURFACE};color:{TEXT};border:1px solid {BORDER};"
            f"border-radius:10px;padding:6px;{FONT_CSS}font-size:12px;}}"
            "QMenu::item{padding:7px 26px 7px 12px;border-radius:7px;}"
            "QMenu::item:selected{background:#23252e;}"
        )
        self.autofile_action = menu.addAction("Auto-file inbox emails")
        self.autofile_action.setCheckable(True)
        self.autofile_action.setToolTip("Suggest a folder when you open an unfiled Inbox email")
        self.setMenu(menu)

    def paintEvent(self, _):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        inset = 1.0 if not self.isDown() else 2.0
        rect = QtCore.QRectF(inset, inset, self.width() - (inset * 2), self.height() - (inset * 2))
        bg = QtGui.QColor("#22242b" if self.isDown() else ("#18191f" if self.underMouse() else "#0d0e12"))
        p.setBrush(bg)
        p.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 40), 1))
        p.drawRoundedRect(rect, self.width() * 0.32, self.height() * 0.32)

        col = QtGui.QColor("#f1f2f6" if self.underMouse() else "#d9dbe3")
        p.setPen(QtGui.QPen(col, 1.6))
        p.setBrush(QtCore.Qt.NoBrush)
        w, h = self.width(), self.height()
        rr = w * 0.15
        p.drawEllipse(QtCore.QPointF(w / 2, h * 0.40), rr, rr)            # head
        shoulders = QtCore.QRectF(w * 0.26, h * 0.58, w * 0.48, h * 0.42)
        p.drawArc(shoulders, 20 * 16, 140 * 16)                          # shoulders
        p.end()


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
    changed = QtCore.Signal()  # emitted when the label (and thus width) changes -> relayout the row

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
        self.changed.emit()


class ApprovalButton(QtWidgets.QPushButton):
    """Separate dropdown: Ask before acting, or act automatically."""
    changed = QtCore.Signal()  # emitted when the label (and thus width) changes -> relayout the row

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
        self.changed.emit()


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
