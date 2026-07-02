"""The Settings panel — a proper popup opened from the profile icon. Consolidates assistant
preferences, brand & colors, and browsing sign-in in one place (replaces the small dropdown menu)."""
from PySide6 import QtCore, QtGui, QtWidgets

from axon.settings import load_settings, save_settings
from axon.ui.theme import PANEL_BG, SURFACE, BORDER, TEXT, MUTED, ACCENT, FONT_CSS

_DEF = {"primary": "1F3A5F", "accent": "E07A2F", "text": "222222", "light": "FFFFFF", "logo": ""}


class SettingsPanel(QtWidgets.QDialog):
    autofile_toggled = QtCore.Signal(bool)
    proactive_toggled = QtCore.Signal(bool)
    signin_requested = QtCore.Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Axon — Settings")
        self.setModal(True)
        self.setMinimumWidth(480)
        self._settings = load_settings()
        self._brand = dict(_DEF)
        self._brand.update(self._settings.get("brand") or {})
        self.setStyleSheet(
            f"QDialog{{background:{PANEL_BG};}}"
            f"QLabel{{color:{TEXT};{FONT_CSS}font-size:12px;}}"
            f"QLineEdit{{background:{SURFACE};color:{TEXT};border:1px solid {BORDER};border-radius:6px;padding:6px;}}"
            f"QCheckBox{{color:{TEXT};font-size:12px;spacing:8px;}}"
            f"QPushButton{{background:{SURFACE};color:{TEXT};border:1px solid {BORDER};border-radius:6px;padding:7px 12px;}}"
            f"QPushButton:hover{{border:1px solid {ACCENT};}}"
        )

        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(22, 18, 22, 18)
        root.setSpacing(6)

        title = QtWidgets.QLabel("Settings")
        title.setStyleSheet(f"color:{TEXT};font-size:18px;font-weight:600;")
        root.addWidget(title)
        root.addSpacing(6)

        # ---- Assistant ----
        self._section(root, "Assistant")
        self._autofile = QtWidgets.QCheckBox("Auto-file inbox emails (suggest a folder when you open an unfiled email)")
        self._autofile.setChecked(bool(self._settings.get("autofile", False)))
        self._autofile.toggled.connect(self._on_autofile)
        root.addWidget(self._autofile)
        self._proactive = QtWidgets.QCheckBox("Proactive nudges (remind me before meetings & about follow-ups)")
        self._proactive.setChecked(self._settings.get("proactive", True))
        self._proactive.toggled.connect(self._on_proactive)
        root.addWidget(self._proactive)

        # ---- Brand & colors ----
        self._section(root, "Brand & colors")
        cap = QtWidgets.QLabel("Used in your PDF reports and branded decks/documents.")
        cap.setStyleSheet(f"color:{MUTED};font-size:11px;")
        root.addWidget(cap)
        self._primary_btn = self._color_row(root, "Primary color", "primary")
        self._accent_btn = self._color_row(root, "Accent color", "accent")
        root.addWidget(QtWidgets.QLabel("Logo (optional)"))
        logo_row = QtWidgets.QHBoxLayout()
        self._logo_edit = QtWidgets.QLineEdit(self._brand.get("logo") or "")
        self._logo_edit.setPlaceholderText("No logo — leave blank")
        browse = QtWidgets.QPushButton("Browse…")
        browse.clicked.connect(self._browse_logo)
        clear = QtWidgets.QPushButton("Clear")
        clear.clicked.connect(self._logo_edit.clear)
        logo_row.addWidget(self._logo_edit, 1)
        logo_row.addWidget(browse)
        logo_row.addWidget(clear)
        root.addLayout(logo_row)

        # ---- Browsing ----
        self._section(root, "Browsing")
        signin = QtWidgets.QPushButton("Sign in to my browser…")
        signin.setToolTip("Open Axon's browser once to sign in to your accounts; reused for all browsing.")
        signin.clicked.connect(self._on_signin)
        signin.setMinimumHeight(34)
        root.addWidget(signin)

        # ---- footer ----
        root.addSpacing(10)
        btns = QtWidgets.QHBoxLayout()
        btns.addStretch(1)
        close = QtWidgets.QPushButton("Close")
        close.clicked.connect(self.reject)
        save = QtWidgets.QPushButton("Save")
        save.setStyleSheet(f"background:{ACCENT};color:white;border:none;border-radius:6px;padding:8px 18px;font-weight:600;")
        save.clicked.connect(self._save)
        btns.addWidget(close)
        btns.addWidget(save)
        root.addLayout(btns)

    # ---- helpers ----
    def _section(self, lay, text):
        lay.addSpacing(12)
        lbl = QtWidgets.QLabel(text.upper())
        lbl.setStyleSheet(f"color:{MUTED};font-size:10px;font-weight:700;letter-spacing:1px;")
        lay.addWidget(lbl)
        line = QtWidgets.QFrame()
        line.setFrameShape(QtWidgets.QFrame.HLine)
        line.setStyleSheet(f"color:{BORDER};background:{BORDER};max-height:1px;")
        lay.addWidget(line)
        lay.addSpacing(6)

    def _color_row(self, lay, label, key):
        row = QtWidgets.QHBoxLayout()
        lbl = QtWidgets.QLabel(label)
        lbl.setFixedWidth(130)
        btn = QtWidgets.QPushButton()
        btn.setFixedHeight(30)
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
        btn.setStyleSheet(f"background:#{hexv};color:{fg};border:1px solid {BORDER};border-radius:6px;font-weight:600;")

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
        s = load_settings(); s["autofile"] = bool(on); save_settings(s)
        self.autofile_toggled.emit(bool(on))

    def _on_proactive(self, on):
        s = load_settings(); s["proactive"] = bool(on); save_settings(s)
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
