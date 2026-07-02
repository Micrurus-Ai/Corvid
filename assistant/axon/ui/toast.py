"""An in-app desktop alert the dot shows itself — styled and animated like Outlook's "new mail"
notification (slides up from the bottom-right, pauses on hover, slides out). Reliable regardless of
Windows notification settings (which drop tray balloons for a pythonw app)."""
from PySide6 import QtCore, QtWidgets

from axon.ui.theme import PANEL_BG, BORDER, TEXT, MUTED, ACCENT, FONT_CSS


class Toast(QtWidgets.QWidget):
    clicked = QtCore.Signal()

    def __init__(self, title, message, action_text=None):
        super().__init__(None, QtCore.Qt.FramelessWindowHint | QtCore.Qt.Tool | QtCore.Qt.WindowStaysOnTopHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setAttribute(QtCore.Qt.WA_ShowWithoutActivating)
        self.setFixedWidth(360)

        card = QtWidgets.QFrame(self)
        card.setObjectName("card")
        card.setStyleSheet(
            f"QFrame#card{{background:{PANEL_BG};border:1px solid {BORDER};border-radius:14px;}}")
        outer = QtWidgets.QHBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(card)

        # left accent stripe (like the coloured bar on a mail alert)
        body = QtWidgets.QHBoxLayout(card)
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(0)
        stripe = QtWidgets.QFrame()
        stripe.setFixedWidth(4)
        stripe.setStyleSheet(f"background:{ACCENT};border-top-left-radius:14px;border-bottom-left-radius:14px;")
        body.addWidget(stripe)

        inner = QtWidgets.QVBoxLayout()
        inner.setContentsMargins(15, 12, 12, 13)
        inner.setSpacing(6)
        body.addLayout(inner, 1)

        top = QtWidgets.QHBoxLayout()
        t = QtWidgets.QLabel(title)
        t.setStyleSheet(f"color:{TEXT};font-weight:700;font-size:12px;{FONT_CSS}")
        close = QtWidgets.QToolButton()
        close.setText("×")
        close.setCursor(QtCore.Qt.PointingHandCursor)
        close.setStyleSheet(f"QToolButton{{color:{MUTED};border:none;font-size:16px;}}QToolButton:hover{{color:{TEXT};}}")
        close.clicked.connect(self._dismiss)
        top.addWidget(t, 1)
        top.addWidget(close)
        inner.addLayout(top)

        m = QtWidgets.QLabel(message)
        m.setWordWrap(True)
        m.setStyleSheet(f"color:#c8cad5;font-size:11.5px;{FONT_CSS}")
        inner.addWidget(m)

        if action_text:
            b = QtWidgets.QPushButton(action_text)
            b.setCursor(QtCore.Qt.PointingHandCursor)
            b.setStyleSheet(f"background:{ACCENT};color:white;border:none;border-radius:8px;"
                            "padding:7px 14px;font-weight:600;")
            b.clicked.connect(self._act)
            row = QtWidgets.QHBoxLayout()
            row.addStretch(1)
            row.addWidget(b)
            inner.addLayout(row)

        self._timer = QtCore.QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self._dismiss)
        self._anim = None

    # pause auto-dismiss while the cursor is over it (like the mail alert)
    def enterEvent(self, e):
        self._timer.stop()
        super().enterEvent(e)

    def leaveEvent(self, e):
        self._timer.start(4000)
        super().leaveEvent(e)

    def _act(self):
        self.clicked.emit()
        self._dismiss()

    def show_at(self, geo):
        self.adjustSize()
        x = geo.right() - self.width() - 22
        end_y = geo.bottom() - self.height() - 48
        start_y = end_y + 46
        self.setWindowOpacity(0.0)
        self.move(x, start_y)
        self.show()
        self.raise_()
        self._anim = QtCore.QParallelAnimationGroup(self)
        pa = QtCore.QPropertyAnimation(self, b"pos")
        pa.setDuration(280)
        pa.setStartValue(QtCore.QPoint(x, start_y))
        pa.setEndValue(QtCore.QPoint(x, end_y))
        pa.setEasingCurve(QtCore.QEasingCurve.OutCubic)
        oa = QtCore.QPropertyAnimation(self, b"windowOpacity")
        oa.setDuration(280)
        oa.setStartValue(0.0)
        oa.setEndValue(1.0)
        self._anim.addAnimation(pa)
        self._anim.addAnimation(oa)
        self._anim.start()
        self._timer.start(9000)   # visible ~9s, like Outlook's new-mail alert

    def _dismiss(self):
        self._timer.stop()
        try:
            here = self.pos()
            self._anim = QtCore.QParallelAnimationGroup(self)
            pa = QtCore.QPropertyAnimation(self, b"pos")
            pa.setDuration(200)
            pa.setStartValue(here)
            pa.setEndValue(QtCore.QPoint(here.x(), here.y() + 46))
            oa = QtCore.QPropertyAnimation(self, b"windowOpacity")
            oa.setDuration(200)
            oa.setStartValue(self.windowOpacity())
            oa.setEndValue(0.0)
            self._anim.addAnimation(pa)
            self._anim.addAnimation(oa)
            self._anim.finished.connect(self.close)
            self._anim.start()
        except Exception:
            self.close()
