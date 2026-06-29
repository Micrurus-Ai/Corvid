"""Screen-region snip selector for screenshots."""
from PySide6 import QtCore, QtGui, QtWidgets
from axon.ui.theme import (ACCENT, ACCENT_2, PANEL_BG, PANEL_BG_2, SURFACE, BORDER, TEXT,
                            MUTED, FONT_FAMILY, FONT_CSS, HEADER_ICON_SIZE, CONTROL_ICON_SIZE)


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
