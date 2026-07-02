"""Settings popup opened from the profile icon."""
from PySide6 import QtCore, QtGui, QtWidgets

from axon.settings import load_settings, save_settings
from axon.ui.theme import PANEL_BG, PANEL_BG_2, SURFACE, SURFACE_2, BORDER, TEXT, MUTED, FONT_CSS

_DEF = {"primary": "1F3A5F", "accent": "E07A2F", "text": "222222", "light": "FFFFFF", "logo": ""}


class SettingsPanel(QtWidgets.QDialog):
    autofile_toggled = QtCore.Signal(bool)
    proactive_toggled = QtCore.Signal(bool)
    signin_requested = QtCore.Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Axon intelligence - Settings")
        self.setModal(True)
        self.setMinimumSize(560, 520)
        self._settings = load_settings()
        self._brand = dict(_DEF)
        self._brand.update(self._settings.get("brand") or {})
        self.setStyleSheet(
            f"QDialog{{background:{PANEL_BG};}}"
            f"QLabel{{color:{TEXT};{FONT_CSS}font-size:12px;}}"
            f"QLabel#title{{font-size:20px;font-weight:700;}}"
            f"QLabel#subtitle{{color:{MUTED};font-size:12px;}}"
            f"QLabel#sectionTitle{{color:{TEXT};font-size:13px;font-weight:700;}}"
            f"QLabel#sectionHint,QLabel#description{{color:{MUTED};font-size:11px;}}"
            f"QFrame#sectionCard{{background:{PANEL_BG_2};border:1px solid {BORDER};border-radius:14px;}}"
            f"QLineEdit{{background:{SURFACE_2};color:{TEXT};border:1px solid {BORDER};border-radius:9px;"
            "padding:9px 11px;selection-background-color:#2f68d8;}}"
            "QLineEdit:focus{border:1px solid #565967;}"
            f"QCheckBox{{color:{TEXT};font-size:12px;font-weight:600;spacing:10px;}}"
            f"QCheckBox::indicator{{width:18px;height:18px;border-radius:5px;border:1px solid #3a3c47;background:{SURFACE};}}"
            "QCheckBox::indicator:hover{border:1px solid #6a6d7d;}"
            "QCheckBox::indicator:checked{background:#f6f7fb;border:1px solid #f6f7fb;}"
            f"QPushButton{{background:{SURFACE};color:{TEXT};border:1px solid {BORDER};border-radius:9px;"
            "padding:9px 14px;font-size:12px;}}"
            "QPushButton:hover{background:#171821;border:1px solid #4a4d5a;}"
            "QPushButton:pressed{background:#0b0c10;}"
            "QPushButton#primaryAction{background:#f6f7fb;color:#050506;border:1px solid #f6f7fb;font-weight:700;}"
            "QPushButton#primaryAction:hover{background:#ffffff;border:1px solid #ffffff;}"
            f"QPushButton#secondaryAction{{background:transparent;color:{TEXT};}}"
            "QPushButton#saveButton{background:#f6f7fb;color:#050506;border:1px solid #f6f7fb;"
            "font-weight:700;padding:9px 20px;}"
            "QPushButton#saveButton:hover{background:#ffffff;border:1px solid #ffffff;}"
        )

        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(22, 20, 22, 18)
        root.setSpacing(14)

        title = QtWidgets.QLabel("Settings")
        title.setObjectName("title")
        root.addWidget(title)
        subtitle = QtWidgets.QLabel("Control automation, branded output, and browser access.")
        subtitle.setObjectName("subtitle")
        root.addWidget(subtitle)

        assistant = self._section(root, "Assistant", "Background actions Axon can help with.")
        self._autofile = self._preference_row(
            assistant,
            "Auto-file inbox emails",
            "Suggest a destination folder when an unfiled email opens.",
        )
        self._autofile.setChecked(bool(self._settings.get("autofile", False)))
        self._autofile.toggled.connect(self._on_autofile)
        self._proactive = self._preference_row(
            assistant,
            "Proactive nudges",
            "Remind you before meetings and about follow-ups.",
        )
        self._proactive.setChecked(self._settings.get("proactive", True))
        self._proactive.toggled.connect(self._on_proactive)

        brand = self._section(root, "Brand and colors", "Used in PDF reports, decks, and documents.")
        self._primary_btn = self._color_row(brand, "Primary", "primary")
        self._accent_btn = self._color_row(brand, "Accent", "accent")
        logo_label = QtWidgets.QLabel("Logo")
        logo_label.setStyleSheet("font-weight:600;")
        brand.addWidget(logo_label)
        logo_row = QtWidgets.QHBoxLayout()
        logo_row.setSpacing(8)
        self._logo_edit = QtWidgets.QLineEdit(self._brand.get("logo") or "")
        self._logo_edit.setPlaceholderText("No logo selected")
        browse = QtWidgets.QPushButton("Browse...")
        browse.clicked.connect(self._browse_logo)
        clear = QtWidgets.QPushButton("Clear")
        clear.clicked.connect(self._logo_edit.clear)
        logo_row.addWidget(self._logo_edit, 1)
        logo_row.addWidget(browse)
        logo_row.addWidget(clear)
        brand.addLayout(logo_row)

        browsing = self._section(root, "Browsing", "Connect the browser once, then reuse it for web tasks.")
        signin = QtWidgets.QPushButton("Sign in to browser")
        signin.setObjectName("primaryAction")
        signin.setToolTip("Open Axon's browser once to sign in to your accounts; reused for all browsing.")
        signin.clicked.connect(self._on_signin)
        signin.setMinimumHeight(40)
        browsing.addWidget(signin)

        root.addStretch(1)
        btns = QtWidgets.QHBoxLayout()
        btns.setSpacing(8)
        btns.addStretch(1)
        close = QtWidgets.QPushButton("Close")
        close.setObjectName("secondaryAction")
        close.clicked.connect(self.reject)
        save = QtWidgets.QPushButton("Save")
        save.setObjectName("saveButton")
        save.clicked.connect(self._save)
        btns.addWidget(close)
        btns.addWidget(save)
        root.addLayout(btns)

    def _section(self, lay, title, hint):
        card = QtWidgets.QFrame()
        card.setObjectName("sectionCard")
        card_lay = QtWidgets.QVBoxLayout(card)
        card_lay.setContentsMargins(16, 14, 16, 14)
        card_lay.setSpacing(10)

        lbl = QtWidgets.QLabel(title)
        lbl.setObjectName("sectionTitle")
        card_lay.addWidget(lbl)

        sub = QtWidgets.QLabel(hint)
        sub.setObjectName("sectionHint")
        sub.setWordWrap(True)
        card_lay.addWidget(sub)

        lay.addWidget(card)
        return card_lay

    def _preference_row(self, lay, title, description):
        wrap = QtWidgets.QFrame()
        row = QtWidgets.QVBoxLayout(wrap)
        row.setContentsMargins(0, 2, 0, 2)
        row.setSpacing(4)
        check = QtWidgets.QCheckBox(title)
        row.addWidget(check)
        desc = QtWidgets.QLabel(description)
        desc.setObjectName("description")
        desc.setWordWrap(True)
        desc.setContentsMargins(28, 0, 0, 0)
        row.addWidget(desc)
        lay.addWidget(wrap)
        return check

    def _color_row(self, lay, label, key):
        row = QtWidgets.QHBoxLayout()
        row.setSpacing(12)
        lbl = QtWidgets.QLabel(label)
        lbl.setStyleSheet("font-weight:600;")
        lbl.setFixedWidth(92)
        btn = QtWidgets.QPushButton()
        btn.setFixedHeight(38)
        btn.setCursor(QtCore.Qt.PointingHandCursor)
        self._paint(btn, key)
        btn.clicked.connect(lambda: self._pick(key, btn))
        row.addWidget(lbl)
        row.addWidget(btn, 1)
        lay.addLayout(row)
        return btn

    def _paint(self, btn, key):
        hexv = (self._brand.get(key) or "888888").lstrip("#")
        try:
            r, g, b = int(hexv[0:2], 16), int(hexv[2:4], 16), int(hexv[4:6], 16)
        except Exception:
            r = g = b = 136
        fg = "#000000" if (r * 0.299 + g * 0.587 + b * 0.114) > 150 else "#FFFFFF"
        btn.setText("#" + hexv.upper())
        btn.setStyleSheet(
            f"background:#{hexv};color:{fg};border:1px solid #3a3c47;"
            "border-radius:10px;font-weight:700;"
        )

    def _pick(self, key, btn):
        cur = QtGui.QColor("#" + (self._brand.get(key) or "888888").lstrip("#"))
        c = QtWidgets.QColorDialog.getColor(cur, self, "Pick " + key + " color")
        if c.isValid():
            self._brand[key] = c.name().lstrip("#").upper()
            self._paint(btn, key)

    def _browse_logo(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Choose a logo image", "", "Images (*.png *.jpg *.jpeg *.gif *.bmp)")
        if path:
            self._logo_edit.setText(path)

    def _on_autofile(self, on):
        s = load_settings()
        s["autofile"] = bool(on)
        save_settings(s)
        self.autofile_toggled.emit(bool(on))

    def _on_proactive(self, on):
        s = load_settings()
        s["proactive"] = bool(on)
        save_settings(s)
        self.proactive_toggled.emit(bool(on))

    def _on_signin(self):
        self.signin_requested.emit()

    def _save(self):
        self._brand["logo"] = self._logo_edit.text().strip()
        s = load_settings()
        b = s.get("brand") or {}
        b.update({k: self._brand[k] for k in ("primary", "accent", "text", "light", "logo")})
        s["brand"] = b
        save_settings(s)
        self.accept()
