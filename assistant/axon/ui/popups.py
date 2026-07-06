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
POPUP_FONT_CSS = "font-family:'Segoe UI', Arial, sans-serif;"


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
        self.resize(520, 600)
        self._eid = None
        self._folders = []

        self.card = PanelFrame(self, fill_color=POPUP_BG, border_color=POPUP_BORDER)
        self._lay = QtWidgets.QVBoxLayout(self.card)
        self._lay.setContentsMargins(24, 22, 24, 20)
        self._lay.setSpacing(7)

        title = QtWidgets.QLabel("Move this email to a folder")
        title.setStyleSheet(f"color:{POPUP_TEXT};{POPUP_FONT_CSS}font-size:22px;font-weight:800;")
        self._lay.addWidget(title)

        self.info = QtWidgets.QLabel("")
        self.info.setWordWrap(True)
        self.info.setStyleSheet(f"color:{POPUP_MUTED};{POPUP_FONT_CSS}font-size:14px;")
        self._lay.addWidget(self.info)

        self.suggestion_area = QtWidgets.QWidget()
        self.suggestion_area.setMinimumHeight(176)
        suggestion_lay = QtWidgets.QVBoxLayout(self.suggestion_area)
        suggestion_lay.setContentsMargins(0, 2, 0, 0)
        suggestion_lay.setSpacing(5)
        suggestion_lay.addWidget(self._section_label("Suggested folders"))
        self._sugg_box = QtWidgets.QVBoxLayout()
        self._sugg_box.setSpacing(5)
        suggestion_lay.addLayout(self._sugg_box)
        suggestion_lay.addStretch(1)
        self._lay.addWidget(self.suggestion_area)

        self._lay.addWidget(self._section_label("Or create a new folder"))
        create_row = QtWidgets.QHBoxLayout()
        create_row.setSpacing(8)
        self.new_folder = QtWidgets.QLineEdit()
        self.new_folder.setStyleSheet(self._field_style())
        create_row.addWidget(self.new_folder, 1)
        create = self._primary_button("Create & Move", min_width=148)
        create.clicked.connect(self._create_and_move)
        create_row.addWidget(create)
        self._lay.addLayout(create_row)

        self._lay.addWidget(self._section_label("Or pick an existing folder"))
        self.search = QtWidgets.QLineEdit()
        self.search.setStyleSheet(self._field_style())
        self.search.textChanged.connect(self._filter_folders)
        self._lay.addWidget(self.search)

        self.folder_list = QtWidgets.QListWidget()
        self.folder_list.setObjectName("FolderPickList")
        self.folder_list.setUniformItemSizes(False)
        self.folder_list.setVerticalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        self.folder_list.setStyleSheet(
            f"QListWidget#FolderPickList{{background:{POPUP_SURFACE};color:{POPUP_TEXT};"
            f"border:1px solid {POPUP_BORDER};border-radius:0;outline:0;{POPUP_FONT_CSS}font-size:13px;}}"
            "QListWidget#FolderPickList::item{padding:0;border:0;}"
            f"QListWidget#FolderPickList::item:selected{{background:#E7E4EE;color:{POPUP_TEXT};}}"
            f"QListWidget#FolderPickList::item:hover{{background:{POPUP_SURFACE_2};}}")
        self.folder_list.itemDoubleClicked.connect(lambda _item: self._move_selected())
        self.folder_list.currentItemChanged.connect(lambda _current, _previous: self._sync_folder_rows())
        self._lay.addWidget(self.folder_list, 1)

        footer = QtWidgets.QHBoxLayout()
        footer.setContentsMargins(0, 10, 0, 0)
        footer.setSpacing(8)
        footer.addStretch(1)
        self.move_other = self._primary_button("Move", min_width=104)
        self.move_other.clicked.connect(self._move_selected)
        footer.addWidget(self.move_other)
        keep = self._secondary_button("Keep in Inbox", min_width=130)
        keep.clicked.connect(self._keep)
        footer.addWidget(keep)
        self._lay.addLayout(footer)
        self._reposition()

    def _section_label(self, text):
        label = QtWidgets.QLabel(text.upper())
        label.setStyleSheet(
            f"color:{POPUP_MUTED};{POPUP_FONT_CSS}font-size:12px;font-weight:800;letter-spacing:0px;")
        return label

    def _field_style(self):
        return (
            f"QLineEdit{{background:{POPUP_SURFACE};color:{POPUP_TEXT};border:1px solid {POPUP_BORDER};"
            f"border-radius:8px;padding:8px 12px;{POPUP_FONT_CSS}font-size:13px;"
            f"selection-background-color:{POPUP_SURFACE_2};}}"
            "QLineEdit:focus{border:1px solid #BBAA94;background:#FFFDF8;}")

    def _primary_button(self, text, min_width=0):
        button = QtWidgets.QPushButton(text)
        button.setMinimumWidth(min_width)
        button.setCursor(QtCore.Qt.PointingHandCursor)
        button.setStyleSheet(
            f"QPushButton{{background:{POPUP_PRIMARY};color:{POPUP_PRIMARY_TEXT};border:1px solid {POPUP_PRIMARY};"
            f"border-radius:0;padding:8px 16px;{POPUP_FONT_CSS}font-size:14px;font-weight:800;}}"
            "QPushButton:hover{background:#302D28;}"
            "QPushButton:pressed{background:#050505;}")
        return button

    def _secondary_button(self, text, min_width=0):
        button = QtWidgets.QPushButton(text)
        button.setMinimumWidth(min_width)
        button.setCursor(QtCore.Qt.PointingHandCursor)
        button.setStyleSheet(
            f"QPushButton{{background:{POPUP_SURFACE};color:{POPUP_TEXT};border:1px solid {POPUP_BORDER};"
            f"border-radius:0;padding:8px 16px;{POPUP_FONT_CSS}font-size:14px;font-weight:500;}}"
            f"QPushButton:hover{{background:{POPUP_SURFACE_2};}}")
        return button

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

    def _folder_label(self, folder):
        name = (folder or "").replace("\\", "/").rstrip("/").split("/")[-1].strip()
        return name or folder or ""

    def _make_folder_item(self, folder):
        item = QtWidgets.QListWidgetItem()
        item.setData(QtCore.Qt.UserRole, folder)
        item.setSizeHint(QtCore.QSize(0, 48))

        row = QtWidgets.QWidget()
        row.setAttribute(QtCore.Qt.WA_TransparentForMouseEvents)
        row.setProperty("folderSelected", False)
        row.setStyleSheet(
            "QWidget{background:transparent;}"
            f"QWidget[folderSelected='true']{{background:#E7E4EE;}}")
        lay = QtWidgets.QVBoxLayout(row)
        lay.setContentsMargins(14, 5, 14, 5)
        lay.setSpacing(0)
        name = QtWidgets.QLabel(self._folder_label(folder))
        name.setStyleSheet(f"color:{POPUP_TEXT};{POPUP_FONT_CSS}font-size:16px;font-weight:800;")
        path = QtWidgets.QLabel(folder)
        path.setStyleSheet(f"color:{POPUP_MUTED};{POPUP_FONT_CSS}font-size:12px;")
        lay.addWidget(name)
        lay.addWidget(path)
        return item, row

    def _filter_folders(self):
        query = self.search.text().strip().lower()
        self.folder_list.clear()
        for folder in self._folders:
            if not query or query in folder.lower():
                item, row = self._make_folder_item(folder)
                self.folder_list.addItem(item)
                self.folder_list.setItemWidget(item, row)
        if self.folder_list.count():
            self.folder_list.setCurrentRow(0)
        self._sync_folder_rows()

    def _sync_folder_rows(self):
        current = self.folder_list.currentItem()
        for i in range(self.folder_list.count()):
            item = self.folder_list.item(i)
            row = self.folder_list.itemWidget(item)
            if not row:
                continue
            selected = item is current
            row.setProperty("folderSelected", selected)
            row.style().unpolish(row)
            row.style().polish(row)

    def present(self, eid, subject, sender, suggestions, folders, screen=None):
        self._eid = eid
        subj = (subject or "(no subject)").strip()
        self.info.setText(f"<span style='font-weight:600'>{subj}</span><br>from {sender or 'unknown'}")
        self._clear_suggestions()
        for name in (suggestions or [])[:3]:
            b = QtWidgets.QPushButton("->  " + self._folder_label(name))
            b.setCursor(QtCore.Qt.PointingHandCursor)
            b.setStyleSheet(
                f"QPushButton{{background:#56545C;color:{POPUP_PRIMARY_TEXT};border:0;border-radius:0;"
                f"padding:7px 18px;text-align:left;{POPUP_FONT_CSS}font-size:14px;font-weight:800;}}"
                "QPushButton:hover{background:#44424A;}")
            b.clicked.connect(lambda _=False, n=name: self._pick(n))
            self._sugg_box.addWidget(b)
        self.new_folder.clear()
        self.search.clear()
        self._folders = list(folders or [])
        self._filter_folders()
        self.resize(520, 600)
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
        self._move_selected()

    def _create_and_move(self):
        folder = self.new_folder.text().strip()
        if folder:
            self._pick(folder)

    def _move_selected(self):
        item = self.folder_list.currentItem()
        folder = item.data(QtCore.Qt.UserRole).strip() if item else ""
        if folder:
            self._pick(folder)

    def _keep(self):
        eid = self._eid
        self.hide()
        if eid:
            self.skipped.emit(eid)
