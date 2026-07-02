"""Guide-mode on-screen overlay: highlight box, panel frame, and screenshot mark view."""
import math

from PySide6 import QtCore, QtGui, QtWidgets
from axon.ui.theme import (ACCENT, ACCENT_2, PANEL_BG, PANEL_BG_2, SURFACE, BORDER, TEXT,
                            MUTED, FONT_FAMILY, FONT_CSS, HEADER_ICON_SIZE, CONTROL_ICON_SIZE)


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
    def __init__(self, parent=None, fill_color=None):
        super().__init__(parent)
        self.setAttribute(QtCore.Qt.WA_StyledBackground, False)
        self._fill_color = fill_color or PANEL_BG

    def paintEvent(self, _):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        rect = QtCore.QRectF(0.5, 0.5, self.width() - 1, self.height() - 1)
        path = QtGui.QPainterPath()
        path.addRoundedRect(rect, 18, 18)

        p.fillPath(path, QtGui.QColor(self._fill_color))

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
