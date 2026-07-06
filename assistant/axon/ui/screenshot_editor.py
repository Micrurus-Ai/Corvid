"""Screenshot editor used by the composer attachment thumbnail."""
from PySide6 import QtCore, QtGui, QtWidgets

from axon.ui.theme import BORDER, FONT_CSS, MUTED, PANEL_BG, PANEL_BG_2, SURFACE, TEXT


def _editor_icon(kind, color="#f7f7fb"):
    pix = QtGui.QPixmap(22, 22)
    pix.fill(QtCore.Qt.transparent)
    painter = QtGui.QPainter(pix)
    painter.setRenderHint(QtGui.QPainter.Antialiasing)
    pen = QtGui.QPen(QtGui.QColor(color), 1.8)
    pen.setCapStyle(QtCore.Qt.RoundCap)
    pen.setJoinStyle(QtCore.Qt.RoundJoin)
    painter.setPen(pen)
    painter.setBrush(QtCore.Qt.NoBrush)

    if kind == "pen":
        painter.drawLine(QtCore.QPointF(6, 16), QtCore.QPointF(15, 7))
        painter.drawLine(QtCore.QPointF(13, 5), QtCore.QPointF(17, 9))
        painter.drawLine(QtCore.QPointF(5, 17), QtCore.QPointF(9, 16))
    elif kind == "highlight":
        painter.drawLine(QtCore.QPointF(6, 15), QtCore.QPointF(14, 7))
        painter.drawLine(QtCore.QPointF(13, 6), QtCore.QPointF(17, 10))
        marker = QtGui.QPen(QtGui.QColor("#ffd84d"), 3.2)
        marker.setCapStyle(QtCore.Qt.RoundCap)
        painter.setPen(marker)
        painter.drawLine(QtCore.QPointF(5, 18), QtCore.QPointF(17, 18))
    elif kind == "box":
        painter.drawRoundedRect(QtCore.QRectF(5, 6, 12, 11), 2, 2)
    elif kind == "arrow":
        painter.drawLine(QtCore.QPointF(6, 16), QtCore.QPointF(16, 6))
        painter.drawLine(QtCore.QPointF(16, 6), QtCore.QPointF(15, 12))
        painter.drawLine(QtCore.QPointF(16, 6), QtCore.QPointF(10, 7))
    elif kind == "text":
        font = QtGui.QFont()
        font.setPointSize(13)
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(QtCore.QRectF(3, 2, 16, 18), QtCore.Qt.AlignCenter, "T")
    elif kind == "redact":
        painter.setPen(QtCore.Qt.NoPen)
        painter.setBrush(QtGui.QColor(color))
        painter.drawRoundedRect(QtCore.QRectF(5, 8, 12, 7), 2, 2)
    elif kind == "undo":
        painter.drawArc(QtCore.QRectF(5, 6, 12, 10), 30 * 16, 260 * 16)
        painter.drawLine(QtCore.QPointF(7, 7), QtCore.QPointF(4, 7))
        painter.drawLine(QtCore.QPointF(7, 7), QtCore.QPointF(7, 4))
    elif kind == "clear":
        painter.drawLine(QtCore.QPointF(7, 7), QtCore.QPointF(15, 15))
        painter.drawLine(QtCore.QPointF(15, 7), QtCore.QPointF(7, 15))
    painter.end()
    return QtGui.QIcon(pix)


class ScreenshotCanvas(QtWidgets.QWidget):
    changed = QtCore.Signal()

    def __init__(self, path, parent=None):
        super().__init__(parent)
        self.setMouseTracking(True)
        self.setMinimumSize(520, 320)
        self._path = path
        self._pix = self._load_pixmap(path)
        self._history = []
        self._tool = "pen"
        self._color = QtGui.QColor("#ff4f4f")
        self._width = 4
        self._zoom = 1.0
        self._last = None
        self._start = None
        self._preview = None
        self.setCursor(QtCore.Qt.CrossCursor)

    def _load_pixmap(self, path):
        image = QtGui.QImage(path)
        if image.isNull():
            return QtGui.QPixmap(path)
        image = image.convertToFormat(QtGui.QImage.Format_ARGB32_Premultiplied)
        flattened = QtGui.QImage(image.size(), QtGui.QImage.Format_ARGB32_Premultiplied)
        flattened.fill(QtGui.QColor("#ffffff"))
        painter = QtGui.QPainter(flattened)
        painter.drawImage(0, 0, image)
        painter.end()
        return QtGui.QPixmap.fromImage(flattened)

    def set_tool(self, tool):
        self._tool = tool
        if tool == "highlight":
            self._color = QtGui.QColor("#ffd84d")

    def set_color(self, color):
        self._color = QtGui.QColor(color)

    def set_width(self, width):
        self._width = int(width)

    def set_zoom(self, value):
        self._zoom = max(0.65, min(1.75, int(value) / 100.0))
        self.update()

    def save(self, path):
        return self._pix.save(path, "PNG")

    def undo(self):
        if self._history:
            self._pix = self._history.pop()
            self.update()
            self.changed.emit()

    def clear_marks(self):
        if self._history:
            self._pix = self._history[0]
            self._history = []
            self.update()
            self.changed.emit()

    def _push_history(self):
        self._history.append(QtGui.QPixmap(self._pix))
        if len(self._history) > 30:
            self._history.pop(0)

    def _image_rect(self):
        if self._pix.isNull():
            return QtCore.QRectF()
        available = QtCore.QSizeF(max(1, self.width() - 40), max(1, self.height() - 40))
        scaled = self._pix.size().scaled(available.toSize(), QtCore.Qt.KeepAspectRatio)
        scaled = QtCore.QSize(int(scaled.width() * self._zoom), int(scaled.height() * self._zoom))
        x = (self.width() - scaled.width()) / 2
        y = (self.height() - scaled.height()) / 2
        return QtCore.QRectF(x, y, scaled.width(), scaled.height())

    def _to_image(self, pos):
        rect = self._image_rect()
        if not rect.contains(pos) or self._pix.isNull():
            return None
        x = (pos.x() - rect.x()) / rect.width() * self._pix.width()
        y = (pos.y() - rect.y()) / rect.height() * self._pix.height()
        return QtCore.QPointF(x, y)

    def _pen(self, highlighter=False):
        if highlighter:
            color = QtGui.QColor("#ffe66d")
            color.setAlpha(56)
        else:
            color = QtGui.QColor(self._color)
        pen = QtGui.QPen(color, self._width * (3 if highlighter else 1))
        pen.setCapStyle(QtCore.Qt.RoundCap)
        pen.setJoinStyle(QtCore.Qt.RoundJoin)
        return pen

    def _draw_line(self, p1, p2, highlighter=False):
        painter = QtGui.QPainter(self._pix)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        if highlighter:
            painter.setCompositionMode(QtGui.QPainter.CompositionMode_Multiply)
        painter.setPen(self._pen(highlighter))
        painter.drawLine(p1, p2)
        painter.end()

    def _draw_shape(self, start, end):
        painter = QtGui.QPainter(self._pix)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        painter.setPen(self._pen(False))
        painter.setBrush(QtCore.Qt.NoBrush)
        if self._tool == "box":
            painter.drawRoundedRect(QtCore.QRectF(start, end).normalized(), 8, 8)
        elif self._tool == "arrow":
            self._paint_arrow(painter, start, end)
        elif self._tool == "redact":
            painter.setPen(QtCore.Qt.NoPen)
            painter.setBrush(QtGui.QColor("#050506"))
            painter.drawRoundedRect(QtCore.QRectF(start, end).normalized(), 6, 6)
        painter.end()

    def _paint_arrow(self, painter, start, end):
        painter.drawLine(start, end)
        line = QtCore.QLineF(start, end)
        if line.length() < 4:
            return
        angle = line.angle()
        head_len = max(10, self._width * 3)
        for delta in (150, -150):
            ray = QtCore.QLineF(end, QtCore.QPointF(end.x() + head_len, end.y()))
            ray.setAngle(angle + delta)
            painter.drawLine(end, ray.p2())

    def _draw_text(self, point):
        text, ok = QtWidgets.QInputDialog.getText(self, "Add text", "Text:")
        if not ok or not text.strip():
            return
        self._push_history()
        painter = QtGui.QPainter(self._pix)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        font = QtGui.QFont()
        font.setPointSize(max(12, self._width * 4))
        font.setBold(True)
        painter.setFont(font)
        painter.setPen(QtGui.QPen(self._color))
        painter.drawText(point, text.strip())
        painter.end()
        self.update()
        self.changed.emit()

    def mousePressEvent(self, event):
        if event.button() != QtCore.Qt.LeftButton:
            return
        pt = self._to_image(event.position())
        if pt is None:
            return
        if self._tool == "text":
            self._draw_text(pt)
            return
        self._push_history()
        self._start = pt
        self._last = pt
        self._preview = None

    def mouseMoveEvent(self, event):
        pt = self._to_image(event.position())
        if pt is None or self._last is None:
            return
        if self._tool in ("pen", "highlight"):
            self._draw_line(self._last, pt, self._tool == "highlight")
            self._last = pt
            self.update()
            self.changed.emit()
        elif self._tool in ("box", "arrow", "redact"):
            self._preview = (self._start, pt)
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() != QtCore.Qt.LeftButton or self._last is None:
            return
        pt = self._to_image(event.position())
        if pt is not None and self._tool in ("box", "arrow", "redact"):
            self._draw_shape(self._start, pt)
            self.changed.emit()
        self._last = None
        self._start = None
        self._preview = None
        self.update()

    def paintEvent(self, _):
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        painter.fillRect(self.rect(), QtGui.QColor("#050506"))
        shell = QtCore.QRectF(8, 8, self.width() - 16, self.height() - 16)
        painter.setBrush(QtGui.QColor("#090a0d"))
        painter.setPen(QtGui.QPen(QtGui.QColor("#11131a"), 1))
        painter.drawRoundedRect(shell, 18, 18)
        rect = self._image_rect()
        if not self._pix.isNull():
            painter.drawPixmap(rect, self._pix, QtCore.QRectF(self._pix.rect()))
        painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 55), 1))
        painter.drawRoundedRect(rect.adjusted(-1, -1, 1, 1), 10, 10)
        if self._preview:
            start, end = self._preview
            sx = rect.x() + start.x() / self._pix.width() * rect.width()
            sy = rect.y() + start.y() / self._pix.height() * rect.height()
            ex = rect.x() + end.x() / self._pix.width() * rect.width()
            ey = rect.y() + end.y() / self._pix.height() * rect.height()
            painter.setPen(self._pen(False))
            if self._tool == "box":
                painter.drawRoundedRect(QtCore.QRectF(QtCore.QPointF(sx, sy), QtCore.QPointF(ex, ey)).normalized(), 8, 8)
            elif self._tool == "arrow":
                self._paint_arrow(painter, QtCore.QPointF(sx, sy), QtCore.QPointF(ex, ey))
            elif self._tool == "redact":
                painter.setPen(QtCore.Qt.NoPen)
                painter.setBrush(QtGui.QColor(5, 5, 6, 210))
                painter.drawRoundedRect(
                    QtCore.QRectF(QtCore.QPointF(sx, sy), QtCore.QPointF(ex, ey)).normalized(), 6, 6)
        painter.end()


class ScreenshotEditor(QtWidgets.QDialog):
    def __init__(self, path, parent=None):
        super().__init__(parent)
        self._path = path
        self.setWindowTitle("Edit screenshot")
        self.setModal(True)
        self.resize(1120, 760)
        self.setMinimumSize(860, 560)
        self.setStyleSheet(
            f"QDialog{{background:{PANEL_BG};}}"
            f"QLabel{{color:{TEXT};{FONT_CSS}}}"
            f"QFrame#toolbar{{background:{PANEL_BG_2};border:1px solid {BORDER};border-radius:14px;}}"
            f"QLabel#toolLabel{{color:{MUTED};font-size:10px;font-weight:700;letter-spacing:1px;}}"
            f"QPushButton{{background:{SURFACE};color:{TEXT};border:1px solid {BORDER};border-radius:9px;"
            "padding:8px 12px;font-size:12px;font-weight:600;}}"
            "QPushButton:hover{background:#181a22;border-color:#4a4d5a;}"
            "QPushButton:checked{background:#f6f7fb;color:#050506;border-color:#f6f7fb;font-weight:800;}"
            f"QPushButton#primary{{background:#f6f7fb;color:#050506;border-color:#f6f7fb;"
            "font-weight:800;border-radius:12px;padding:10px 18px;}}"
            f"QPushButton#primary:hover{{background:#ffffff;border-color:#ffffff;}}"
            f"QPushButton#ghost{{background:#101014;color:{TEXT};border-color:{BORDER};"
            "border-radius:12px;padding:10px 18px;}}"
            f"QPushButton#ghost:hover{{background:#171821;border-color:#4a4d5a;}}"
            f"QSlider::groove:horizontal{{height:4px;background:{BORDER};border-radius:2px;}}"
            "QSlider::handle:horizontal{width:16px;margin:-6px 0;border-radius:8px;background:#f6f7fb;}"
        )

        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(18, 16, 18, 16)
        root.setSpacing(12)

        self.canvas = ScreenshotCanvas(path)

        header = QtWidgets.QHBoxLayout()
        title = QtWidgets.QLabel("Edit screenshot")
        title.setStyleSheet("font-size:18px;font-weight:700;")
        header.addWidget(title)
        hint = QtWidgets.QLabel("Markup")
        hint.setStyleSheet(f"color:{MUTED};font-size:12px;")
        header.addWidget(hint)
        header.addStretch(1)
        root.addLayout(header)

        toolbar = QtWidgets.QFrame()
        toolbar.setObjectName("toolbar")
        tools = QtWidgets.QHBoxLayout(toolbar)
        tools.setContentsMargins(10, 10, 10, 10)
        tools.setSpacing(8)
        tools.addWidget(self._caption("TOOLS"))
        self._tool_group = QtWidgets.QButtonGroup(self)
        self._tool_group.setExclusive(True)
        for label, tool in (
            ("Pen", "pen"),
            ("Highlight", "highlight"),
            ("Box", "box"),
            ("Arrow", "arrow"),
            ("Text", "text"),
            ("Redact", "redact"),
        ):
            btn = self._icon_button(tool, label)
            btn.setCheckable(True)
            if tool == "pen":
                btn.setChecked(True)
            btn.clicked.connect(lambda _=False, t=tool: self.canvas.set_tool(t))
            self._tool_group.addButton(btn)
            tools.addWidget(btn)

        tools.addSpacing(8)
        tools.addWidget(self._caption("COLOR"))
        self._color_buttons = []
        for color in ("#ff4f4f", "#ffd84d", "#4da3ff", "#ffffff", "#111111"):
            swatch = QtWidgets.QPushButton("")
            swatch.setFixedSize(28, 28)
            swatch.setStyleSheet(
                f"QPushButton{{background:{color};border:1px solid #5a5d68;border-radius:14px;padding:0;}}"
                "QPushButton:hover{border:2px solid #ffffff;}"
            )
            swatch.clicked.connect(lambda _=False, c=color: self.canvas.set_color(c))
            self._color_buttons.append(swatch)
            tools.addWidget(swatch)

        tools.addSpacing(8)
        tools.addWidget(self._caption("SIZE"))
        width = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        width.setRange(2, 12)
        width.setValue(4)
        width.setFixedWidth(88)
        width.valueChanged.connect(self.canvas_width_changed)
        tools.addWidget(width)

        tools.addSpacing(8)
        tools.addWidget(self._caption("ZOOM"))
        zoom = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        zoom.setRange(65, 175)
        zoom.setValue(100)
        zoom.setFixedWidth(92)
        zoom.valueChanged.connect(self.canvas.set_zoom)
        tools.addWidget(zoom)
        tools.addStretch(1)

        undo = self._icon_button("undo", "Undo")
        undo.clicked.connect(lambda: self.canvas.undo())
        clear = self._icon_button("clear", "Clear")
        clear.clicked.connect(lambda: self.canvas.clear_marks())
        tools.addWidget(undo)
        tools.addWidget(clear)
        root.addWidget(toolbar)

        width.valueChanged.connect(self.canvas.set_width)
        root.addWidget(self.canvas, 1)

        footer = QtWidgets.QHBoxLayout()
        footer.addStretch(1)
        cancel = QtWidgets.QPushButton("Cancel")
        cancel.setObjectName("ghost")
        cancel.clicked.connect(self.reject)
        save = QtWidgets.QPushButton("Save edits")
        save.setObjectName("primary")
        save.clicked.connect(self._save)
        footer.addWidget(cancel)
        footer.addWidget(save)
        root.addLayout(footer)

    def _caption(self, text):
        label = QtWidgets.QLabel(text)
        label.setObjectName("toolLabel")
        return label

    def _icon_button(self, icon_name, tooltip):
        btn = QtWidgets.QPushButton("")
        btn.setIcon(_editor_icon(icon_name))
        btn.setIconSize(QtCore.QSize(22, 22))
        btn.setFixedSize(42, 38)
        btn.setToolTip(tooltip)
        btn.setAccessibleName(tooltip)
        return btn

    def canvas_width_changed(self, value):
        if hasattr(self, "canvas"):
            self.canvas.set_width(value)

    def _save(self):
        if not self.canvas.save(self._path):
            QtWidgets.QMessageBox.warning(self, "Edit screenshot", "Could not save the edited screenshot.")
            return
        self.accept()
