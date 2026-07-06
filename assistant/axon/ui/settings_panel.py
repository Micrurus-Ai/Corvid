"""Settings popup opened from the profile icon. Visual design follows the Figma
'Axon Settings UI' spec: dark page, a segmented tab bar with a raised active pill,
section cards (icon + title + divider + rows), green toggles, and a white Save button."""
import threading

from PySide6 import QtCore, QtGui, QtWidgets

from axon.settings import load_settings, save_settings
from axon.outlook.tone import my_tone, save_tone, learn_my_tone
from axon import archive_config

# --- palette (matches the Figma design tokens) -------------------------------
PAGE = "#0d0e11"          # dialog background
CARD = "#121318"          # section card fill
CARD_BORDER = "#212430"
DIVIDER = "#20222b"
ICON_BG = "#1b1c22"
ICON_BORDER = "#2c2e37"
T_TITLE = "#f4f4f8"
T_DESC = "#84848f"
T_LABEL = "#e6e7ec"
T_FOOT = "#7c7f8b"
FIELD_BG = "#191a20"
FIELD_BORDER = "#2a2c35"
FIELD_FOCUS = "#4a4d5a"
NAV_BG = "#101116"
NAV_BORDER = "#212430"
TAB_ACTIVE = "#2b2c34"
GREEN = "#34c759"
WHITE_BTN = "#f5f6f8"
FONT = "font-family:'Segoe UI','Inter',Arial,sans-serif;"

_DEF = {"primary": "1F3A5F", "accent": "E07A2F", "text": "222222", "light": "FFFFFF", "logo": ""}


class SwitchCheckBox(QtWidgets.QCheckBox):
    """iOS-style toggle: green when on, grey track + white knob when off."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self.setFixedSize(44, 24)

    def hitButton(self, pos):
        return self.rect().contains(pos)

    def paintEvent(self, _event):
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        track = QtCore.QRectF(1, 2, 42, 20)
        checked = self.isChecked()
        painter.setPen(QtCore.Qt.NoPen)
        painter.setBrush(QtGui.QColor(GREEN if checked else "#3a3c44"))
        painter.drawRoundedRect(track, 10, 10)
        knob_x = 23 if checked else 3
        painter.setBrush(QtGui.QColor("#ffffff"))
        painter.drawEllipse(QtCore.QRectF(knob_x, 4, 16, 16))


class Segmented(QtWidgets.QFrame):
    """A rounded segmented control (e.g. Save folders: Both / Email / Attachments)."""

    def __init__(self, options, current=0, parent=None):
        super().__init__(parent)
        self.setObjectName("segmented")
        self._buttons = []
        row = QtWidgets.QHBoxLayout(self)
        row.setContentsMargins(3, 3, 3, 3)
        row.setSpacing(3)
        group = QtWidgets.QButtonGroup(self)
        group.setExclusive(True)
        for i, text in enumerate(options):
            b = QtWidgets.QPushButton(text)
            b.setObjectName("segment")
            b.setCheckable(True)
            b.setCursor(QtCore.Qt.PointingHandCursor)
            b.setMinimumHeight(34)
            b.setChecked(i == current)
            group.addButton(b, i)
            row.addWidget(b, 1)
            self._buttons.append(b)
        self.setStyleSheet(
            f"QFrame#segmented{{background:{FIELD_BG};border:1px solid {FIELD_BORDER};border-radius:11px;}}"
            f"QPushButton#segment{{background:transparent;color:{T_DESC};border:none;border-radius:8px;"
            f"padding:6px 10px;font-size:12px;font-weight:600;{FONT}}}"
            "QPushButton#segment:hover{color:#ffffff;}"
            f"QPushButton#segment:checked{{background:{TAB_ACTIVE};color:#ffffff;}}"
        )

    def current_index(self):
        for i, b in enumerate(self._buttons):
            if b.isChecked():
                return i
        return 0


class SettingsPanel(QtWidgets.QDialog):
    autofile_toggled = QtCore.Signal(bool)
    proactive_toggled = QtCore.Signal(bool)
    signin_requested = QtCore.Signal()
    tone_learned = QtCore.Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Axon intelligence — Settings")
        self.setModal(True)
        self.setMinimumSize(760, 660)
        self._settings = load_settings()
        self._brand = dict(_DEF)
        self._brand.update(self._settings.get("brand") or {})
        self._color_edits = {}
        self._color_swatches = {}

        self.setStyleSheet(
            f"QDialog{{background:{PAGE};}}"
            f"QToolTip{{background:{WHITE_BTN};color:#050506;border:1px solid {CARD_BORDER};"
            f"border-radius:8px;padding:6px 8px;font-size:12px;{FONT}}}"
            f"QLabel{{color:{T_TITLE};{FONT}font-size:12px;}}"
            f"QLabel#title{{color:#ffffff;font-size:19px;font-weight:800;}}"
            f"QLabel#subtitle{{color:{T_DESC};font-size:12.5px;}}"
            f"QLabel#sectionTitle{{color:{T_TITLE};font-size:13.5px;font-weight:700;}}"
            f"QLabel#sectionHint{{color:{T_DESC};font-size:11.5px;}}"
            f"QLabel#prefTitle{{color:{T_TITLE};font-size:13px;font-weight:600;}}"
            f"QLabel#prefDesc{{color:{T_DESC};font-size:11.5px;}}"
            f"QLabel#fieldLabel{{color:{T_LABEL};font-size:12px;font-weight:600;}}"
            f"QLabel#fieldHint{{color:{T_DESC};font-size:11px;}}"
            f"QLabel#description{{color:{T_DESC};font-size:11.5px;}}"
            f"QLabel#footnote{{color:{T_FOOT};font-size:11.5px;}}"
            f"QLabel#sectionIcon{{background:{ICON_BG};color:#cfd0d8;border:1px solid {ICON_BORDER};"
            "border-radius:8px;font-size:13px;font-weight:700;}"
            f"QFrame#sectionCard{{background:{CARD};border:1px solid {CARD_BORDER};border-radius:14px;}}"
            f"QFrame#divider{{background:{DIVIDER};border:none;}}"
            f"QFrame#footerBar{{background:transparent;border-top:1px solid {DIVIDER};}}"
            "QScrollArea{background:transparent;border:none;}"
            f"QScrollArea > QWidget > QWidget{{background:{PAGE};}}"
            "QScrollBar:vertical{background:transparent;width:9px;margin:4px 2px;}"
            f"QScrollBar::handle:vertical{{background:{FIELD_BORDER};border-radius:4px;min-height:40px;}}"
            "QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical{height:0;}"
            "QScrollBar::add-page:vertical,QScrollBar::sub-page:vertical{background:transparent;}"
            f"QLineEdit{{background:{FIELD_BG};color:{T_TITLE};border:1px solid {FIELD_BORDER};border-radius:10px;"
            f"padding:10px 12px;font-size:12.5px;{FONT}selection-background-color:#2f68d8;}}"
            f"QLineEdit:focus{{border:1px solid {FIELD_FOCUS};}}"
            f"QPlainTextEdit{{background:{FIELD_BG};color:{T_TITLE};border:1px solid {FIELD_BORDER};"
            f"border-radius:11px;padding:12px;font-size:12.5px;{FONT}}}"
            f"QPlainTextEdit:focus{{border:1px solid {FIELD_FOCUS};}}"
        )

        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # --- header -----------------------------------------------------------
        head = QtWidgets.QVBoxLayout()
        head.setContentsMargins(26, 22, 26, 16)
        head.setSpacing(4)
        title = QtWidgets.QLabel("Settings")
        title.setObjectName("title")
        head.addWidget(title)
        subtitle = QtWidgets.QLabel("Automation, writing style, brand assets, and archive rules.")
        subtitle.setObjectName("subtitle")
        head.addWidget(subtitle)

        # segmented tab bar
        nav = QtWidgets.QFrame()
        nav.setObjectName("navBar")
        nav.setStyleSheet(
            f"QFrame#navBar{{background:{NAV_BG};border:1px solid {NAV_BORDER};border-radius:12px;}}"
            f"QPushButton#navTab{{background:transparent;color:{T_DESC};border:none;border-radius:9px;"
            f"padding:11px 18px;font-size:12px;font-weight:600;{FONT}}}"
            "QPushButton#navTab:hover:!checked{color:#eaeaf0;}"
            f"QPushButton#navTab:checked{{background:{TAB_ACTIVE};color:#ffffff;}}"
        )
        nav_row = QtWidgets.QHBoxLayout(nav)
        nav_row.setContentsMargins(5, 5, 5, 5)
        nav_row.setSpacing(4)
        self._nav_group = QtWidgets.QButtonGroup(self)
        self._nav_group.setExclusive(True)
        for i, name in enumerate(("General", "Writing", "Brand", "Email archive")):
            b = QtWidgets.QPushButton(name)
            b.setObjectName("navTab")
            b.setCheckable(True)
            b.setCursor(QtCore.Qt.PointingHandCursor)
            b.setMinimumHeight(44)
            b.setChecked(i == 0)
            self._nav_group.addButton(b, i)
            nav_row.addWidget(b)
        nav_row.addStretch(1)
        self._nav_group.idClicked.connect(lambda idx: self._stack.setCurrentIndex(idx))
        head_wrap = QtWidgets.QHBoxLayout()
        head_wrap.setContentsMargins(0, 8, 0, 0)
        head_wrap.addWidget(nav)
        head_wrap.addStretch(1)
        head.addLayout(head_wrap)
        root.addLayout(head)

        top_line = QtWidgets.QFrame()
        top_line.setObjectName("divider")
        top_line.setFixedHeight(1)
        root.addWidget(top_line)

        # --- stacked pages ----------------------------------------------------
        self._stack = QtWidgets.QStackedWidget()
        root.addWidget(self._stack, 1)
        general = self._page()
        writing = self._page()
        brand_page = self._page()
        archive_page = self._page()

        self._build_general(general)
        self._build_writing(writing)
        self._build_brand(brand_page)
        self._build_archive(archive_page)

        # --- footer -----------------------------------------------------------
        footer = QtWidgets.QFrame()
        footer.setObjectName("footerBar")
        frow = QtWidgets.QHBoxLayout(footer)
        frow.setContentsMargins(26, 16, 26, 16)
        frow.setSpacing(10)
        note = QtWidgets.QLabel("Changes apply after saving.")
        note.setObjectName("footnote")
        frow.addWidget(note)
        frow.addStretch(1)
        close = QtWidgets.QPushButton("Close")
        close.setCursor(QtCore.Qt.PointingHandCursor)
        close.setStyleSheet(
            f"QPushButton{{background:transparent;color:{T_LABEL};border:none;border-radius:10px;"
            f"padding:10px 18px;font-size:12.5px;font-weight:600;{FONT}}}"
            "QPushButton:hover{color:#ffffff;}"
        )
        self._describe_button(close, "Close Settings without saving unsaved edits.")
        close.clicked.connect(self.reject)
        save = QtWidgets.QPushButton("Save")
        save.setCursor(QtCore.Qt.PointingHandCursor)
        save.setMinimumSize(64, 40)
        save.setStyleSheet(
            f"QPushButton{{background:{WHITE_BTN};color:#050506;border:none;border-radius:12px;"
            f"padding:10px 22px;font-size:12.5px;font-weight:800;{FONT}}}"
            "QPushButton:hover{background:#ffffff;}"
            "QPushButton:pressed{background:#e2e4ea;}"
        )
        self._describe_button(save, "Save writing tone, brand, and archive settings, then close.")
        save.clicked.connect(self._save)
        frow.addWidget(close)
        frow.addWidget(save)
        root.addWidget(footer)

    # --- page / section scaffolding ------------------------------------------
    def _page(self):
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        body = QtWidgets.QWidget()
        lay = QtWidgets.QVBoxLayout(body)
        lay.setContentsMargins(24, 20, 24, 16)
        lay.setSpacing(16)
        scroll.setWidget(body)
        self._stack.addWidget(scroll)
        return lay

    def _section(self, lay, title, hint, icon="•"):
        card = QtWidgets.QFrame()
        card.setObjectName("sectionCard")
        card_lay = QtWidgets.QVBoxLayout(card)
        card_lay.setContentsMargins(18, 16, 18, 18)
        card_lay.setSpacing(14)

        header = QtWidgets.QHBoxLayout()
        header.setSpacing(12)
        ic = QtWidgets.QLabel(icon)
        ic.setObjectName("sectionIcon")
        ic.setAlignment(QtCore.Qt.AlignCenter)
        ic.setFixedSize(28, 28)
        header.addWidget(ic, 0, QtCore.Qt.AlignTop)
        copy = QtWidgets.QVBoxLayout()
        copy.setSpacing(2)
        lbl = QtWidgets.QLabel(title)
        lbl.setObjectName("sectionTitle")
        copy.addWidget(lbl)
        sub = QtWidgets.QLabel(hint)
        sub.setObjectName("sectionHint")
        sub.setWordWrap(True)
        copy.addWidget(sub)
        header.addLayout(copy, 1)
        card_lay.addLayout(header)
        card_lay.addWidget(self._divider())
        lay.addWidget(card)
        return card_lay

    def _divider(self):
        line = QtWidgets.QFrame()
        line.setObjectName("divider")
        line.setFixedHeight(1)
        return line

    def _preference_row(self, title, description):
        wrap = QtWidgets.QWidget()
        row = QtWidgets.QHBoxLayout(wrap)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(14)
        copy = QtWidgets.QVBoxLayout()
        copy.setSpacing(3)
        label = QtWidgets.QLabel(title)
        label.setObjectName("prefTitle")
        copy.addWidget(label)
        desc = QtWidgets.QLabel(description)
        desc.setObjectName("prefDesc")
        desc.setWordWrap(True)
        copy.addWidget(desc)
        row.addLayout(copy, 1)
        check = SwitchCheckBox()
        check.setToolTip(description)
        check.setAccessibleName(title)
        row.addWidget(check, 0, QtCore.Qt.AlignVCenter)
        return wrap, check

    def _field_label(self, lay, title, hint=None):
        row = QtWidgets.QHBoxLayout()
        row.setContentsMargins(0, 2, 0, 0)
        row.setSpacing(8)
        label = QtWidgets.QLabel(title)
        label.setObjectName("fieldLabel")
        row.addWidget(label)
        if hint:
            hint_label = QtWidgets.QLabel(hint)
            hint_label.setObjectName("fieldHint")
            row.addWidget(hint_label)
        row.addStretch(1)
        lay.addLayout(row)

    def _description(self, text):
        label = QtWidgets.QLabel(text)
        label.setObjectName("description")
        label.setWordWrap(True)
        return label

    def _describe_button(self, button, description):
        button.setToolTip(description)
        button.setAccessibleName(button.text())
        button.setAccessibleDescription(description)

    def _ghost_button(self, text):
        b = QtWidgets.QPushButton(text)
        b.setCursor(QtCore.Qt.PointingHandCursor)
        b.setMinimumHeight(40)
        b.setStyleSheet(
            f"QPushButton{{background:{FIELD_BG};color:{T_LABEL};border:1px solid {FIELD_BORDER};"
            f"border-radius:10px;padding:9px 16px;font-size:12px;font-weight:600;{FONT}}}"
            "QPushButton:hover{background:#22242c;border-color:#3a3d47;color:#ffffff;}"
            "QPushButton:pressed{background:#141519;}"
            f"QPushButton:disabled{{background:#151619;color:{T_FOOT};border-color:{FIELD_BORDER};}}"
        )
        return b

    # --- General --------------------------------------------------------------
    def _build_general(self, lay):
        assistant = self._section(lay, "Assistant", "How Axon works alongside you in the inbox.", "✦")
        af_wrap, self._autofile = self._preference_row(
            "Auto-file inbox emails",
            "Suggests a destination folder whenever an unfiled email is opened.")
        self._autofile.setChecked(bool(self._settings.get("autofile", False)))
        self._autofile.toggled.connect(self._on_autofile)
        assistant.addWidget(af_wrap)
        assistant.addWidget(self._divider())
        pn_wrap, self._proactive = self._preference_row(
            "Proactive nudges",
            "Reminds you before meetings and about follow-ups that are going quiet.")
        self._proactive.setChecked(self._settings.get("proactive", True))
        self._proactive.toggled.connect(self._on_proactive)
        assistant.addWidget(pn_wrap)

        browsing = self._section(lay, "Browsing", "Web access for quotes, lookups, and portals.", "◍")
        b_row = QtWidgets.QHBoxLayout()
        b_row.setSpacing(14)
        copy = QtWidgets.QVBoxLayout()
        copy.setSpacing(3)
        bt = QtWidgets.QLabel("Browser session")
        bt.setObjectName("prefTitle")
        copy.addWidget(bt)
        bd = QtWidgets.QLabel("Opens Axon's browser profile so web tasks reuse your signed-in sessions.")
        bd.setObjectName("prefDesc")
        bd.setWordWrap(True)
        copy.addWidget(bd)
        status = QtWidgets.QLabel("●  Not signed in")
        status.setStyleSheet(f"color:{T_DESC};font-size:11px;")
        copy.addWidget(status)
        b_row.addLayout(copy, 1)
        signin = self._ghost_button("Sign in to browser")
        self._describe_button(
            signin, "Open Axon's browser once to sign in to your accounts; reused for all browsing tasks.")
        signin.clicked.connect(self._on_signin)
        b_row.addWidget(signin, 0, QtCore.Qt.AlignVCenter)
        browsing.addLayout(b_row)
        lay.addStretch(1)

    # --- Writing --------------------------------------------------------------
    def _build_writing(self, lay):
        tone = self._section(lay, "Writing tone", "How Axon writes every email draft.", "✎")
        self.tone_learned.connect(self._on_tone_learned)
        self._tone_edit = QtWidgets.QPlainTextEdit(my_tone())
        self._tone_edit.setPlaceholderText(
            "Friendly and concise. Greet with the first name. Keep paragraphs short. "
            "Sign off with 'Best regards, Disan'. No jargon, no filler phrases.")
        self._tone_edit.setMinimumHeight(160)
        tone.addWidget(self._tone_edit)

        row = QtWidgets.QHBoxLayout()
        row.setSpacing(14)
        self._learn_btn = self._ghost_button("Learn from my Sent emails")
        self._describe_button(
            self._learn_btn, "Analyse your recent Sent items and fill in your writing style automatically.")
        self._learn_btn.clicked.connect(self._learn_tone)
        row.addWidget(self._learn_btn, 0, QtCore.Qt.AlignTop)
        self._tone_status = QtWidgets.QLabel(
            "Reads your recent Sent mail and drafts a tone profile you can review before saving.")
        self._tone_status.setObjectName("description")
        self._tone_status.setWordWrap(True)
        row.addWidget(self._tone_status, 1)
        tone.addLayout(row)
        lay.addStretch(1)

    # --- Brand ----------------------------------------------------------------
    def _build_brand(self, lay):
        brand = self._section(lay, "Brand and colors",
                              "Applied to covers, headers, and branded exports.", "◈")
        cols = QtWidgets.QHBoxLayout()
        cols.setSpacing(16)
        cols.addWidget(self._color_column(
            "Primary", "primary", "Main headers, covers, and prominent branded elements."), 1)
        cols.addWidget(self._color_column(
            "Accent", "accent", "Highlights, chart details, and supporting marks."), 1)
        brand.addLayout(cols)
        brand.addWidget(self._divider())

        self._field_label(brand, "Logo")
        logo_row = QtWidgets.QHBoxLayout()
        logo_row.setSpacing(8)
        self._logo_edit = QtWidgets.QLineEdit(self._brand.get("logo") or "")
        self._logo_edit.setPlaceholderText("No logo selected")
        browse = self._ghost_button("Browse…")
        self._describe_button(browse, "Choose a logo image file for branded exports.")
        browse.clicked.connect(self._browse_logo)
        clear = self._ghost_button("Clear")
        self._describe_button(clear, "Remove the selected logo path.")
        clear.clicked.connect(self._logo_edit.clear)
        logo_row.addWidget(self._logo_edit, 1)
        logo_row.addWidget(browse)
        logo_row.addWidget(clear)
        brand.addLayout(logo_row)
        brand.addWidget(self._description(
            "PNG, JPG, JPEG, GIF or BMP. Used on branded exports and covers."))
        lay.addStretch(1)

    def _color_column(self, label, key, desc):
        col = QtWidgets.QWidget()
        v = QtWidgets.QVBoxLayout(col)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(8)
        lbl = QtWidgets.QLabel(label)
        lbl.setObjectName("fieldLabel")
        v.addWidget(lbl)
        row = QtWidgets.QHBoxLayout()
        row.setSpacing(10)
        swatch = QtWidgets.QPushButton()
        swatch.setFixedSize(36, 36)
        swatch.setCursor(QtCore.Qt.PointingHandCursor)
        self._describe_button(swatch, f"Pick the {label.lower()} brand color.")
        swatch.clicked.connect(lambda: self._pick(key))
        edit = QtWidgets.QLineEdit()
        edit.setMaxLength(7)
        edit.editingFinished.connect(lambda: self._sync_from_edit(key))
        self._color_edits[key] = edit
        self._color_swatches[key] = swatch
        self._paint_swatch(key)
        row.addWidget(swatch)
        row.addWidget(edit, 1)
        v.addLayout(row)
        v.addWidget(self._description(desc))
        return col

    def _paint_swatch(self, key):
        hexv = (self._brand.get(key) or "888888").lstrip("#").upper()
        self._color_swatches[key].setStyleSheet(
            f"background:#{hexv};border:1px solid {ICON_BORDER};border-radius:10px;")
        edit = self._color_edits[key]
        if edit.text().lstrip("#").upper() != hexv:
            edit.setText("#" + hexv)

    def _sync_from_edit(self, key):
        v = self._color_edits[key].text().strip().lstrip("#")
        if len(v) == 6:
            try:
                int(v, 16)
                self._brand[key] = v.upper()
            except ValueError:
                pass
        self._paint_swatch(key)

    def _pick(self, key):
        cur = QtGui.QColor("#" + (self._brand.get(key) or "888888").lstrip("#"))
        c = QtWidgets.QColorDialog.getColor(cur, self, "Pick " + key + " color")
        if c.isValid():
            self._brand[key] = c.name().lstrip("#").upper()
            self._paint_swatch(key)

    def _browse_logo(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Choose a logo image", "", "Images (*.png *.jpg *.jpeg *.gif *.bmp)")
        if path:
            self._logo_edit.setText(path)

    # --- Email archive --------------------------------------------------------
    def _build_archive(self, lay):
        self._arch = archive_config.load()
        loc = self._section(lay, "Archive locations",
                            "Where archived emails and attachments are stored.", "▤")
        self._arch_client = QtWidgets.QLineEdit(self._arch.get("client_base", ""))
        self._arch_client.setPlaceholderText(r"T:\Archive\Clients   (paste with a sample code)")
        self._field_label(loc, "Client base path")
        loc.addLayout(self._path_row(self._arch_client))
        loc.addWidget(self._description("Folder used as the starting point for client email archives."))

        self._arch_supplier = QtWidgets.QLineEdit(self._arch.get("supplier_base", ""))
        self._arch_supplier.setPlaceholderText("Same as client base path")
        self._field_label(loc, "Supplier base path", "Optional")
        loc.addLayout(self._path_row(self._arch_supplier))
        loc.addWidget(self._description("Leave blank to reuse the client base path."))

        self._field_label(loc, "Default subfolder", "Optional")
        self._arch_sub = QtWidgets.QLineEdit(self._arch.get("default_subfolder", ""))
        self._arch_sub.setPlaceholderText("e.g. Quotes")
        loc.addWidget(self._arch_sub)
        loc.addWidget(self._description("Preselected when saving archived content."))

        rules = self._section(lay, "Archive rules",
                             "How Axon names and saves archived items.", "≡")
        self._field_label(rules, "Country to code")
        self._arch_codes = QtWidgets.QPlainTextEdit(
            "\n".join(f"{k}={v}" for k, v in (self._arch.get("country_codes") or {}).items()))
        self._arch_codes.setPlaceholderText("Belgium=AB\nFrance=FR\nGermany=DE")
        self._arch_codes.setMinimumHeight(96)
        rules.addWidget(self._arch_codes)
        rules.addWidget(self._description("One mapping per line, e.g. Belgium=AB."))

        self._field_label(rules, "Save folders")
        _sm = (self._arch.get("save_mode") or "both").lower()
        cur = {"both": 0, "email": 1, "attachments": 2}.get(_sm, 0)
        self._arch_save = Segmented(
            ["Both (email + attachments)", "Email (.msg)", "Attachments only"], cur)
        rules.addWidget(self._arch_save)
        lay.addStretch(1)

    def _path_row(self, edit):
        row = QtWidgets.QHBoxLayout()
        row.setSpacing(8)
        row.addWidget(edit, 1)
        b = self._ghost_button("Browse…")
        self._describe_button(b, "Choose this folder from your computer.")
        b.clicked.connect(lambda: self._browse_into(edit))
        row.addWidget(b)
        return row

    def _browse_into(self, edit):
        d = QtWidgets.QFileDialog.getExistingDirectory(self, "Choose folder", edit.text() or "")
        if d:
            edit.setText(d.replace("/", "\\"))

    # --- persistence + handlers ----------------------------------------------
    def _save_archive(self):
        codes = {}
        for line in self._arch_codes.toPlainText().splitlines():
            if "=" in line:
                k, v = line.split("=", 1)
                k, v = k.strip(), v.strip()
                if k and v:
                    codes[k] = v
        archive_config.save({
            "client_base": self._arch_client.text().strip(),
            "supplier_base": self._arch_supplier.text().strip(),
            "country_codes": codes,
            "save_mode": ["both", "email", "attachments"][self._arch_save.current_index()],
            "default_subfolder": self._arch_sub.text().strip(),
        })

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

    def _learn_tone(self):
        """Analyse the user's Sent items (off the UI thread) and fill in the tone box."""
        self._learn_btn.setEnabled(False)
        self._tone_status.setText("Reading your Sent items…")

        def work():
            guide = ""
            try:
                learn_my_tone({})   # reads Sent, derives a style guide, saves it
                guide = my_tone()
            except Exception:
                guide = ""
            self.tone_learned.emit(guide)
        threading.Thread(target=work, daemon=True).start()

    @QtCore.Slot(str)
    def _on_tone_learned(self, guide):
        if guide:
            self._tone_edit.setPlainText(guide)
            self._tone_status.setText("Learned from your Sent items — review and Save.")
        else:
            self._tone_status.setText("Couldn't read your Sent items (is Outlook set up?).")
        self._learn_btn.setEnabled(True)

    def _save(self):
        for key in ("primary", "accent"):
            self._sync_from_edit(key)
        save_tone(self._tone_edit.toPlainText())   # persist the writing tone
        self._save_archive()                        # persist the email-archive config
        self._brand["logo"] = self._logo_edit.text().strip()
        s = load_settings()
        b = s.get("brand") or {}
        b.update({k: self._brand[k] for k in ("primary", "accent", "text", "light", "logo")})
        s["brand"] = b
        save_settings(s)
        self.accept()
