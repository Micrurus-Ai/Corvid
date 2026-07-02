"""Frameless popups: action approval and inbox folder-pick."""
import datetime

from PySide6 import QtCore, QtWidgets

from axon.ui.highlight import PanelFrame
from axon.ui.theme import FONT_CSS

POPUP_BG = "#F7F1E8"
POPUP_SURFACE = "#FFF9F0"
POPUP_SURFACE_2 = "#EFE4D4"
POPUP_BORDER = "#D8CBB9"
POPUP_TEXT = "#191611"
POPUP_MUTED = "#70685D"
POPUP_PRIMARY = "#111111"
POPUP_PRIMARY_TEXT = "#FFF9F0"


def _resolve_preset(i):
    n = datetime.datetime.now()
    if i == 0:
        return n + datetime.timedelta(hours=1)
    if i == 1:
        e = n.replace(hour=18, minute=0, second=0, microsecond=0)
        return e if e > n else n + datetime.timedelta(hours=1)
    if i == 2:
        return (n + datetime.timedelta(days=1)).replace(hour=8, minute=0, second=0, microsecond=0)
    if i == 3:
        return (n + datetime.timedelta(days=3)).replace(hour=9, minute=0, second=0, microsecond=0)
    return n + datetime.timedelta(hours=1)


class WhenDialog(QtWidgets.QDialog):
    """Pick when to send: one-click presets + a custom date/time. Returns .when (datetime)."""
    _PRESETS = ["In 1 hour", "This evening (6 PM)", "Tomorrow (8 AM)", "In 3 days"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.when = None
        self.remind = False
        self.setWindowTitle("Send later")
        self.setModal(True)
        self.setStyleSheet(
            f"QDialog{{background:{POPUP_BG};}}"
            f"QLabel{{color:{POPUP_TEXT};{FONT_CSS}font-size:12px;}}"
            f"QPushButton{{background:{POPUP_SURFACE};color:{POPUP_TEXT};border:1px solid {POPUP_BORDER};"
            "border-radius:9px;padding:9px 12px;text-align:left;font-weight:600;}"
            f"QPushButton:hover{{background:{POPUP_SURFACE_2};border:1px solid #C6B69F;}}"
            f"QDateTimeEdit{{background:{POPUP_SURFACE};color:{POPUP_TEXT};border:1px solid {POPUP_BORDER};"
            "border-radius:9px;padding:7px;}}"
            f"QCheckBox{{color:{POPUP_TEXT};{FONT_CSS}font-size:12px;spacing:8px;}}")
        lay = QtWidgets.QVBoxLayout(self)
        lay.setContentsMargins(16, 14, 16, 14)
        lay.setSpacing(8)
        lay.addWidget(QtWidgets.QLabel("When should Axon send this email?"))
        for i, label in enumerate(self._PRESETS):
            b = QtWidgets.QPushButton("   " + label)
            b.setCursor(QtCore.Qt.PointingHandCursor)
            b.clicked.connect(lambda _=False, idx=i: self._pick_preset(idx))
            lay.addWidget(b)
        row = QtWidgets.QHBoxLayout()
        row.addWidget(QtWidgets.QLabel("Or:"))
        self._dt = QtWidgets.QDateTimeEdit(QtCore.QDateTime.currentDateTime().addSecs(3600))
        self._dt.setCalendarPopup(True)
        self._dt.setDisplayFormat("ddd dd MMM yyyy  HH:mm")
        row.addWidget(self._dt, 1)
        setb = QtWidgets.QPushButton("Set")
        setb.setStyleSheet("text-align:center;")
        setb.clicked.connect(self._pick_custom)
        row.addWidget(setb)
        lay.addLayout(row)
        self._remind_cb = QtWidgets.QCheckBox("Remind me ~10 min before it sends")
        lay.addWidget(self._remind_cb)

    def _pick_preset(self, i):
        self.when = _resolve_preset(i)
        self.remind = self._remind_cb.isChecked()
        self.accept()

    def _pick_custom(self):
        self.when = self._dt.dateTime().toPython()
        self.remind = self._remind_cb.isChecked()
        self.accept()


class ApprovalPopup(QtWidgets.QWidget):
    """Small always-on-top prompt asking the user to Approve / Skip / (for emails) Send Later."""
    decided = QtCore.Signal(object)   # True = now, False = skip, ISO string = send later

    def __init__(self):
        super().__init__()
        self.setWindowFlags(
            QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.Tool
        )
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.resize(450, 172)
        self.card = PanelFrame(self, fill_color=POPUP_BG, border_color=POPUP_BORDER)
        lay = QtWidgets.QVBoxLayout(self.card)
        lay.setContentsMargins(20, 18, 20, 16)
        lay.setSpacing(10)
        title = QtWidgets.QLabel("Approve this action?")
        title.setStyleSheet(f"color:{POPUP_TEXT};{FONT_CSS}font-size:15px;font-weight:700;")
        lay.addWidget(title)
        self.msg = QtWidgets.QLabel("")
        self.msg.setWordWrap(True)
        self.msg.setStyleSheet(f"color:{POPUP_TEXT};{FONT_CSS}font-size:13px;")
        lay.addWidget(self.msg, 1)
        row = QtWidgets.QHBoxLayout()
        row.addStretch(1)
        skip = QtWidgets.QPushButton("Skip")
        self._later_btn = QtWidgets.QPushButton("Send Later")
        approve = QtWidgets.QPushButton("Approve")
        for b in (skip, self._later_btn, approve):
            b.setCursor(QtCore.Qt.PointingHandCursor)
        _ghost = (f"QPushButton{{background:{POPUP_SURFACE};color:{POPUP_TEXT};border:1px solid {POPUP_BORDER};"
                  f"border-radius:10px;padding:8px 16px;{FONT_CSS}font-size:13px;font-weight:600;}}"
                  f"QPushButton:hover{{background:{POPUP_SURFACE_2};border:1px solid #C6B69F;}}")
        skip.setStyleSheet(_ghost)
        self._later_btn.setStyleSheet(_ghost)
        approve.setStyleSheet(
            f"QPushButton{{background:{POPUP_PRIMARY};color:{POPUP_PRIMARY_TEXT};border:none;"
            f"border-radius:10px;padding:8px 20px;{FONT_CSS}font-size:13px;font-weight:700;}}"
            "QPushButton:hover{background:#2a2721;}")
        skip.clicked.connect(lambda: self._decide(False))
        self._later_btn.clicked.connect(self._pick_later)
        approve.clicked.connect(lambda: self._decide(True))
        row.addWidget(skip)
        row.addWidget(self._later_btn)
        row.addWidget(approve)
        lay.addLayout(row)
        self._reposition()

    def _reposition(self):
        self.card.setGeometry(8, 8, self.width() - 16, self.height() - 16)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._reposition()

    def ask(self, description, screen=None, allow_later=False):
        self.msg.setText(description)
        self._later_btn.setVisible(bool(allow_later))   # only email sends can be deferred
        self.adjustSize()
        self.resize(max(470, self.width()), self.height())
        geo = (screen.availableGeometry() if screen is not None
               else QtWidgets.QApplication.primaryScreen().availableGeometry())
        self.move(geo.center().x() - self.width() // 2, geo.top() + 120)
        self.show()
        self.raise_()
        self.activateWindow()

    def _pick_later(self):
        dlg = WhenDialog(self)
        if dlg.exec() == QtWidgets.QDialog.Accepted and dlg.when is not None:
            self.hide()
            self.decided.emit({"send_at": dlg.when.strftime("%Y-%m-%dT%H:%M:%S"),
                               "remind": bool(dlg.remind)})

    def _decide(self, decision):
        self.hide()
        self.decided.emit(decision)


class FolderPickPopup(QtWidgets.QWidget):
    """When you open an unfiled Inbox email, this offers the best-matching subfolders to move it to.
    Picking one moves it; 'Keep in Inbox' dismisses."""
    chosen = QtCore.Signal(str, str)   # entry_id, folder
    skipped = QtCore.Signal(str)       # entry_id

    def __init__(self):
        super().__init__()
        self.setWindowFlags(
            QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.Tool)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.resize(440, 250)
        self._eid = None
        self.card = PanelFrame(self, fill_color=POPUP_BG, border_color=POPUP_BORDER)
        self._lay = QtWidgets.QVBoxLayout(self.card)
        self._lay.setContentsMargins(20, 18, 20, 16)
        self._lay.setSpacing(10)

        title = QtWidgets.QLabel("File this email?")
        title.setStyleSheet(f"color:{POPUP_TEXT};{FONT_CSS}font-size:15px;font-weight:700;")
        self._lay.addWidget(title)
        self.info = QtWidgets.QLabel("")
        self.info.setWordWrap(True)
        self.info.setStyleSheet(f"color:{POPUP_MUTED};{FONT_CSS}font-size:12px;")
        self._lay.addWidget(self.info)

        self._sugg_box = QtWidgets.QVBoxLayout()
        self._sugg_box.setSpacing(6)
        self._lay.addLayout(self._sugg_box)

        row = QtWidgets.QHBoxLayout()
        row.setSpacing(8)
        self.other = QtWidgets.QComboBox()
        self.other.setStyleSheet(
            f"QComboBox{{background:{POPUP_SURFACE};color:{POPUP_TEXT};border:1px solid {POPUP_BORDER};border-radius:9px;"
            f"padding:6px 10px;{FONT_CSS}font-size:12px;}}"
            "QComboBox::drop-down{border:0;width:18px;}"
            f"QComboBox QAbstractItemView{{background:{POPUP_SURFACE};color:{POPUP_TEXT};"
            f"selection-background-color:{POPUP_SURFACE_2};selection-color:{POPUP_TEXT};}}")
        row.addWidget(self.other, 1)
        self.move_other = QtWidgets.QPushButton("Move")
        self.move_other.setCursor(QtCore.Qt.PointingHandCursor)
        self.move_other.setStyleSheet(
            f"QPushButton{{background:{POPUP_PRIMARY};color:{POPUP_PRIMARY_TEXT};border:1px solid {POPUP_PRIMARY};"
            f"border-radius:9px;padding:7px 14px;{FONT_CSS}font-size:12px;font-weight:700;}}"
            "QPushButton:hover{background:#2a2721;}")
        self.move_other.clicked.connect(self._move_other)
        row.addWidget(self.move_other)
        self._lay.addLayout(row)

        keep_row = QtWidgets.QHBoxLayout()
        keep_row.addStretch(1)
        keep = QtWidgets.QPushButton("Keep in Inbox")
        keep.setCursor(QtCore.Qt.PointingHandCursor)
        keep.setStyleSheet(
            f"QPushButton{{background:transparent;color:{POPUP_MUTED};border:none;{FONT_CSS}font-size:12px;}}"
            f"QPushButton:hover{{color:{POPUP_TEXT};}}")
        keep.clicked.connect(self._keep)
        keep_row.addWidget(keep)
        self._lay.addLayout(keep_row)
        self._reposition()

    def _reposition(self):
        self.card.setGeometry(8, 8, self.width() - 16, self.height() - 16)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._reposition()

    def _clear_suggestions(self):
        while self._sugg_box.count():
            it = self._sugg_box.takeAt(0)
            w = it.widget()
            if w:
                w.deleteLater()

    def present(self, eid, subject, sender, suggestions, folders, screen=None):
        self._eid = eid
        subj = (subject or "(no subject)").strip()
        self.info.setText(f"<b>{subj}</b><br>from {sender or 'unknown'}")
        self._clear_suggestions()
        for name in (suggestions or [])[:3]:
            b = QtWidgets.QPushButton("→  " + name)
            b.setCursor(QtCore.Qt.PointingHandCursor)
            b.setStyleSheet(
                f"QPushButton{{background:{POPUP_SURFACE};color:{POPUP_TEXT};border:1px solid {POPUP_BORDER};border-radius:9px;"
                f"padding:8px 12px;text-align:left;{FONT_CSS}font-size:13px;font-weight:700;}}"
                f"QPushButton:hover{{background:{POPUP_SURFACE_2};}}")
            b.clicked.connect(lambda _=False, n=name: self._pick(n))
            self._sugg_box.addWidget(b)
        self.other.clear()
        self.other.addItems(folders or [])
        self.adjustSize()
        self.resize(max(440, self.width()), max(250, self.sizeHint().height()))
        geo = (screen.availableGeometry() if screen is not None
               else QtWidgets.QApplication.primaryScreen().availableGeometry())
        self.move(geo.right() - self.width() - 40, geo.top() + 80)
        self.show()
        self.raise_()
        self.activateWindow()

    def _pick(self, folder):
        eid = self._eid
        self.hide()
        if eid:
            self.chosen.emit(eid, folder)

    def _move_other(self):
        folder = self.other.currentText().strip()
        if folder:
            self._pick(folder)

    def _keep(self):
        eid = self._eid
        self.hide()
        if eid:
            self.skipped.emit(eid)
