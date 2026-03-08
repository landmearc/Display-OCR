import sys
import os
import re
import csv
import json
import ctypes
import tempfile
import time
import shutil
from pathlib import Path
from ctypes import wintypes

from PySide6.QtCore import Qt, QRect, Signal, QTimer, QSize
from PySide6.QtGui import (
    QAction, QColor, QCursor, QGuiApplication, QIcon, QPainter, QPen, QPixmap
)
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QCheckBox, QSpinBox, QMessageBox, QDialog,
    QTextEdit, QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView, QSystemTrayIcon, QMenu
)

import pyperclip
import pyautogui
import pytesseract

# ---------- Windows DPI warning mitigation ----------
# Qt より先に DPI awareness を設定しておくことで
# "SetProcessDpiAwarenessContext() failed" 警告を抑制する
try:
    ctypes.windll.shcore.SetProcessDpiAwarenessContext(-4)  # PER_MONITOR_AWARE_V2
except Exception:
    pass  # 既に設定済みの場合は無視

os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "0")
os.environ.setdefault("QT_SCALE_FACTOR_ROUNDING_POLICY", "PassThrough")
os.environ.setdefault("QT_LOGGING_RULES", "qt.qpa.window=false")

# ---------- Paths ----------
APP_DIR = Path.home() / ".display_ocr"
APP_DIR.mkdir(parents=True, exist_ok=True)
SETTINGS_FILE = APP_DIR / "settings.json"
HISTORY_FILE = APP_DIR / "history.json"
MASTER_TERMS_FILE = APP_DIR / "master_terms.csv"
REPLACE_FILE = APP_DIR / "ocr_replace.csv"
STARTUP_CMD = Path(os.environ.get("APPDATA", "")) / "Microsoft/Windows/Start Menu/Programs/Startup/DisplayOCR.cmd"

DEFAULT_SETTINGS = {
    "trim_spaces": True,
    "normalize_chiseki": True,
    "convert_chome": True,
    "fullwidth_alnum": False,
    "keep_chiseki_halfwidth": True,
    "history_limit": 50,
    "startup": False,
    "window_x": None,
    "window_y": None,
}

KANJI_DIGITS = {0: "〇", 1: "一", 2: "二", 3: "三", 4: "四", 5: "五", 6: "六", 7: "七", 8: "八", 9: "九"}
OCR_LANG = "jpn+jpn_vert+eng"

# ---------- Windows paste helpers ----------
user32 = ctypes.windll.user32

INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
VK_CONTROL = 0x11
VK_V = 0x56


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class INPUTUNION(ctypes.Union):
    _fields_ = [("ki", KEYBDINPUT)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("union", INPUTUNION)]


def _keyboard_input(vk: int, flags: int = 0) -> INPUT:
    return INPUT(
        type=INPUT_KEYBOARD,
        union=INPUTUNION(
            ki=KEYBDINPUT(wVk=vk, wScan=0, dwFlags=flags, time=0, dwExtraInfo=None)
        ),
    )


def send_key_combo(*keys: int):
    inputs = []
    for key in keys:
        inputs.append(_keyboard_input(key, 0))
    for key in reversed(keys):
        inputs.append(_keyboard_input(key, KEYEVENTF_KEYUP))
    arr = (INPUT * len(inputs))(*inputs)
    user32.SendInput(len(inputs), ctypes.byref(arr), ctypes.sizeof(INPUT))


def send_ctrl_v():
    send_key_combo(VK_CONTROL, VK_V)




# ---------- Runtime helpers ----------
def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def get_runtime_root() -> Path:
    if is_frozen():
        return Path(getattr(sys, "_MEIPASS"))
    return Path(__file__).resolve().parent


def get_app_launch_target() -> Path:
    if is_frozen():
        return Path(sys.executable).resolve()
    return Path(__file__).resolve()


def get_bundled_tesseract_dir() -> Path:
    return get_runtime_root() / "tesseract"


def get_bundled_tesseract_exe() -> Path:
    return get_bundled_tesseract_dir() / "tesseract.exe"


def get_bundled_tessdata_dir() -> Path:
    return get_bundled_tesseract_dir() / "tessdata"


def locate_tesseract():
    # 1) bundled in PyInstaller package
    bundled_exe = get_bundled_tesseract_exe()
    if bundled_exe.exists():
        return bundled_exe, get_bundled_tessdata_dir()

    # 2) common install paths
    for candidate in [
        Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
        Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
    ]:
        if candidate.exists():
            return candidate, candidate.parent / "tessdata"

    # 3) PATH
    found = shutil.which("tesseract")
    if found:
        exe = Path(found)
        return exe, exe.parent / "tessdata"

    return None, None


def configure_tesseract():
    exe_path, tessdata_dir = locate_tesseract()
    if not exe_path:
        return None, None

    pytesseract.pytesseract.tesseract_cmd = str(exe_path)

    if tessdata_dir and tessdata_dir.exists():
        os.environ["TESSDATA_PREFIX"] = str(tessdata_dir)

    return exe_path, tessdata_dir


# ---------- Data helpers ----------
def load_json(path: Path, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def save_json(path: Path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


class Settings:
    def __init__(self):
        self.data = DEFAULT_SETTINGS.copy()
        self.data.update(load_json(SETTINGS_FILE, {}))

    def save(self):
        save_json(SETTINGS_FILE, self.data)

    def __getitem__(self, item):
        return self.data[item]

    def __setitem__(self, key, value):
        self.data[key] = value


class HistoryStore:
    def __init__(self, settings: Settings):
        self.settings = settings
        self.items = load_json(HISTORY_FILE, [])
        if not isinstance(self.items, list):
            self.items = []
        self.trim()

    def trim(self):
        limit = max(1, int(self.settings["history_limit"]))
        self.items = self.items[:limit]

    def save(self):
        self.trim()
        save_json(HISTORY_FILE, self.items)

    def add(self, text: str, source: str = "ocr"):
        text = (text or "").strip()
        if not text:
            return
        self.items.insert(0, {"text": text, "source": source})
        self.trim()
        self.save()

    def update_text(self, index: int, text: str):
        if 0 <= index < len(self.items):
            self.items[index]["text"] = text
            self.save()

    def delete(self, index: int):
        if 0 <= index < len(self.items):
            del self.items[index]
            self.save()


# ---------- CSV helpers ----------
def ensure_master_csv():
    if MASTER_TERMS_FILE.exists():
        return
    with open(MASTER_TERMS_FILE, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["text"])
        writer.writerow(["長野県上伊那郡辰野町"])
        writer.writerow(["宅地"])


def load_master_terms():
    ensure_master_csv()
    result = []
    try:
        with open(MASTER_TERMS_FILE, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            first = True
            for row in reader:
                if not row:
                    continue
                val = row[0].strip()
                if first and val.lower() == "text":
                    first = False
                    continue
                first = False
                if val:
                    result.append(val)
    except Exception:
        pass
    return result


def save_master_terms(rows):
    with open(MASTER_TERMS_FILE, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["text"])
        for text in rows:
            text = str(text).strip()
            if text:
                writer.writerow([text])


def ensure_replace_csv():
    if REPLACE_FILE.exists():
        return
    with open(REPLACE_FILE, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["wrong", "correct", "count"])
        writer.writeheader()


def load_replace_rules():
    ensure_replace_csv()
    rules = []
    try:
        with open(REPLACE_FILE, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                wrong = (row.get("wrong") or "").strip()
                correct = (row.get("correct") or "").strip()
                count = int((row.get("count") or "0").strip() or 0)
                if wrong:
                    rules.append((wrong, correct, count))
    except Exception:
        pass
    rules.sort(key=lambda x: (-x[2], -len(x[0]), x[0]))
    return rules


def apply_replace_rules(text: str):
    result = text
    for wrong, correct, _count in load_replace_rules():
        if wrong:
            result = result.replace(wrong, correct)
    return result


def learn_replace_rule(old_text: str, new_text: str):
    old_text = (old_text or "").strip()
    new_text = (new_text or "").strip()
    if not old_text or not new_text or old_text == new_text:
        return
    if len(old_text) > 80 or len(new_text) > 80:
        return

    ensure_replace_csv()
    with open(REPLACE_FILE, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))

    found = False
    for row in rows:
        if (row.get("wrong") or "").strip() == old_text and (row.get("correct") or "").strip() == new_text:
            c = int((row.get("count") or "0").strip() or 0)
            row["count"] = str(c + 1)
            found = True
            break

    if not found:
        rows.append({"wrong": old_text, "correct": new_text, "count": "1"})

    with open(REPLACE_FILE, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["wrong", "correct", "count"])
        writer.writeheader()
        writer.writerows(rows)


# ---------- Text normalization ----------
def trim_spaces(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def int_to_kanji(num: int) -> str:
    if num == 0:
        return KANJI_DIGITS[0]
    parts = []
    units = [(1000, "千"), (100, "百"), (10, "十")]
    n = num
    for base, unit in units:
        d = n // base
        n %= base
        if d:
            if d > 1:
                parts.append(KANJI_DIGITS[d])
            parts.append(unit)
    if n:
        parts.append(KANJI_DIGITS[n])
    return "".join(parts)


def convert_chome(text: str) -> str:
    def repl(m):
        n = int(m.group(1))
        return f"{int_to_kanji(n)}丁目"
    return re.sub(r"(\d+)丁目", repl, text or "")


def normalize_chiseki(text: str) -> str:
    s = (text or "").strip()
    table = str.maketrans({
        "０": "0", "１": "1", "２": "2", "３": "3", "４": "4",
        "５": "5", "６": "6", "７": "7", "８": "8", "９": "9",
        "：": ":", "．": ".", "。": ".", "　": " ",
    })
    s = s.translate(table)
    s = re.sub(r"\s+", "", s)
    s = s.replace("㎡", "")

    m = re.fullmatch(r"(\d+):(\d*)", s)
    if m:
        int_part = m.group(1)
        dec_part = m.group(2)
        if dec_part == "":
            return int_part
        dec_part = (dec_part + "00")[:2]
        return f"{int_part}.{dec_part}"
    return s


def to_fullwidth_alnum(text: str) -> str:
    out = []
    for c in text or "":
        code = ord(c)
        if 0x30 <= code <= 0x39 or 0x41 <= code <= 0x5A or 0x61 <= code <= 0x7A:
            out.append(chr(code + 0xFEE0))
        elif c in "-–—ｰ":
            out.append("－")
        elif c == ".":
            out.append("．")
        else:
            out.append(c)
    return "".join(out)


def should_treat_as_chiseki(raw_text: str) -> bool:
    s = trim_spaces(raw_text)
    s = s.translate(str.maketrans({
        "０": "0", "１": "1", "２": "2", "３": "3", "４": "4",
        "５": "5", "６": "6", "７": "7", "８": "8", "９": "9", "：": ":"
    }))
    return bool(re.fullmatch(r"\d+:(\d*)", s))


def normalize_text(text: str, settings: Settings):
    raw_text = text or ""
    result = raw_text
    chiseki_hit = False

    if settings["trim_spaces"]:
        result = trim_spaces(result)

    result = apply_replace_rules(result)

    if settings["normalize_chiseki"] and should_treat_as_chiseki(result):
        result = normalize_chiseki(result)
        chiseki_hit = True

    if settings["convert_chome"]:
        result = convert_chome(result)

    if settings["fullwidth_alnum"]:
        if not (chiseki_hit and settings["keep_chiseki_halfwidth"]):
            result = to_fullwidth_alnum(result)

    return result, raw_text, chiseki_hit


# ---------- OCR overlay ----------
class CaptureOverlay(QWidget):
    captured = Signal(str)
    canceled = Signal()

    def __init__(self, settings: Settings):
        super().__init__()
        self.settings = settings
        self.start_pos = None
        self.end_pos = None
        self.current_screen = None
        self._continuous = False
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setCursor(Qt.CrossCursor)
        self.hide()

    def begin_capture(self, continuous: bool = False):
        self._continuous = continuous
        cursor_pos = QCursor.pos()
        screen = QGuiApplication.screenAt(cursor_pos) or QGuiApplication.primaryScreen()
        self.current_screen = screen
        geo = screen.geometry()
        self.setGeometry(geo)
        self.start_pos = None
        self.end_pos = None
        self.show()
        self.raise_()
        self.activateWindow()
        self.update()

    def paintEvent(self, _event):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor(0, 0, 0, 80))
        if self.start_pos and self.end_pos:
            rect = QRect(self.start_pos, self.end_pos).normalized()
            p.setCompositionMode(QPainter.CompositionMode_Clear)
            p.fillRect(rect, Qt.transparent)
            p.setCompositionMode(QPainter.CompositionMode_SourceOver)
            p.setPen(QPen(QColor(0, 180, 255), 2))
            p.drawRect(rect)
        if self._continuous:
            hint = "  連続スキャン中  ／  ESC で終了  "
            p.setCompositionMode(QPainter.CompositionMode_SourceOver)
            font = p.font()
            font.setPointSize(13)
            font.setBold(True)
            p.setFont(font)
            fm = p.fontMetrics()
            text_w = fm.horizontalAdvance(hint)
            text_h = fm.height()
            pad_x, pad_y = 16, 8
            banner_w = text_w + pad_x * 2
            banner_h = text_h + pad_y * 2
            cx = self.rect().center().x()
            by = self.rect().bottom() - banner_h - 20
            banner_rect = QRect(cx - banner_w // 2, by, banner_w, banner_h)
            p.setBrush(QColor(30, 30, 30, 210))
            p.setPen(Qt.NoPen)
            p.drawRoundedRect(banner_rect, 6, 6)
            p.setPen(QColor(255, 220, 60))
            p.drawText(banner_rect, Qt.AlignCenter, hint)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.start_pos = event.position().toPoint()
            self.end_pos = self.start_pos
            self.update()

    def mouseMoveEvent(self, event):
        if self.start_pos:
            self.end_pos = event.position().toPoint()
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self.start_pos and self.end_pos:
            rect = QRect(self.start_pos, self.end_pos).normalized()
            self.hide()
            QTimer.singleShot(80, lambda: self.capture_and_ocr(rect))

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.hide()
            self.canceled.emit()

    def capture_and_ocr(self, local_rect: QRect):
        if self.current_screen is None:
            self.canceled.emit()
            return
        x, y, w, h = local_rect.x(), local_rect.y(), local_rect.width(), local_rect.height()
        if w < 3 or h < 3:
            self.canceled.emit()
            return

        pixmap = self.current_screen.grabWindow(0, x, y, w, h)
        if pixmap.isNull():
            self.canceled.emit()
            return

        tmp_path = Path(tempfile.gettempdir()) / "display_ocr_capture.png"
        pixmap.save(str(tmp_path), "PNG")

        exe_path, tessdata_dir = configure_tesseract()
        if not exe_path:
            QMessageBox.warning(
                self,
                "Tesseract未検出",
                "同梱版または設定済みの Tesseract が見つかりませんでした。\n"
                "配布版なら同梱漏れ、開発版なら Tesseract の設定を確認してください。",
            )
            self.canceled.emit()
            return

        try:
            config_parts = ["--oem", "3", "--psm", "6"]
            if tessdata_dir and tessdata_dir.exists():
                config_parts += ["--tessdata-dir", str(tessdata_dir)]
            config = " ".join(f'"{x}"' if " " in x else x for x in config_parts)

            text = pytesseract.image_to_string(str(tmp_path), lang=OCR_LANG, config=config)
        except Exception as e:
            QMessageBox.warning(
                self,
                "OCRエラー",
                f"OCRに失敗しました。\n\n{e}"
            )
            
            self.canceled.emit()
            return

        self.captured.emit((text or "").strip())


# ---------- Dialogs ----------
class HistoryDialog(QDialog):
    def __init__(self, history: HistoryStore, parent=None):
        super().__init__(parent)
        self.history = history
        self.setWindowTitle("履歴表示・編集")
        self.resize(700, 450)

        lay = QVBoxLayout(self)
        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["No", "文字列"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        lay.addWidget(self.table)

        btns = QHBoxLayout()
        self.btn_save = QPushButton("保存")
        self.btn_delete = QPushButton("選択行削除")
        self.btn_close = QPushButton("閉じる")
        btns.addWidget(self.btn_save)
        btns.addWidget(self.btn_delete)
        btns.addStretch(1)
        btns.addWidget(self.btn_close)
        lay.addLayout(btns)

        self.btn_save.clicked.connect(self.save_rows)
        self.btn_delete.clicked.connect(self.delete_selected)
        self.btn_close.clicked.connect(self.accept)

        self.reload()

    def reload(self):
        self.table.setRowCount(len(self.history.items))
        for i, row in enumerate(self.history.items):
            no_item = QTableWidgetItem(str(i + 1))
            no_item.setFlags(no_item.flags() & ~Qt.ItemIsEditable)
            txt_item = QTableWidgetItem(row.get("text", ""))
            self.table.setItem(i, 0, no_item)
            self.table.setItem(i, 1, txt_item)

    def save_rows(self):
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 1)
            self.history.update_text(i, item.text() if item else "")
        QMessageBox.information(self, "履歴", "保存しました。")

    def delete_selected(self):
        rows = sorted({idx.row() for idx in self.table.selectionModel().selectedRows()}, reverse=True)
        for r in rows:
            self.history.delete(r)
        self.reload()


class MasterTermsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("定型文字列編集")
        self.resize(500, 420)
        lay = QVBoxLayout(self)

        self.table = QTableWidget(0, 1)
        self.table.setHorizontalHeaderLabels(["定型文字列"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        lay.addWidget(self.table)

        row1 = QHBoxLayout()
        self.btn_add = QPushButton("行追加")
        self.btn_del = QPushButton("選択行削除")
        self.btn_open = QPushButton("CSVを開く")
        row1.addWidget(self.btn_add)
        row1.addWidget(self.btn_del)
        row1.addWidget(self.btn_open)
        row1.addStretch(1)
        lay.addLayout(row1)

        row2 = QHBoxLayout()
        self.btn_save = QPushButton("保存")
        self.btn_close = QPushButton("閉じる")
        row2.addWidget(self.btn_save)
        row2.addStretch(1)
        row2.addWidget(self.btn_close)
        lay.addLayout(row2)

        self.btn_add.clicked.connect(self.add_row)
        self.btn_del.clicked.connect(self.delete_row)
        self.btn_open.clicked.connect(self.open_csv)
        self.btn_save.clicked.connect(self.save_rows)
        self.btn_close.clicked.connect(self.accept)
        self.reload()

    def reload(self):
        rows = load_master_terms()
        self.table.setRowCount(len(rows))
        for i, text in enumerate(rows):
            self.table.setItem(i, 0, QTableWidgetItem(text))

    def add_row(self):
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QTableWidgetItem(""))

    def delete_row(self):
        rows = sorted({idx.row() for idx in self.table.selectionModel().selectedRows()}, reverse=True)
        for r in rows:
            self.table.removeRow(r)

    def open_csv(self):
        os.startfile(str(MASTER_TERMS_FILE))

    def save_rows(self):
        rows = []
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            rows.append(item.text() if item else "")
        save_master_terms(rows)
        QMessageBox.information(self, "定型文字列", "保存しました。")


class PasteDialog(QDialog):
    def __init__(self, main_window, current_text="", raw_ocr="", continuous_mode=False):
        super().__init__(main_window)
        self.main_window = main_window
        self.raw_ocr = raw_ocr or current_text
        self.continuous_mode = continuous_mode
        self.wants_next_scan = False
        self.setWindowTitle("ペースト")
        self.resize(760, 560)

        lay = QVBoxLayout(self)

        if continuous_mode:
            banner = QLabel("🔄  連続スキャン中  ―  「続けてスキャン」で次へ、「スキャン終了」で終了")
            banner.setStyleSheet(
                "background-color: #1a5fa8; color: white; font-weight: bold;"
                "padding: 6px 10px; border-radius: 4px;"
            )
            banner.setAlignment(Qt.AlignCenter)
            lay.addWidget(banner)

        lay.addWidget(QLabel("OCR結果（修正可）"))
        self.edit = QTextEdit()
        self.edit.setPlainText(current_text)
        lay.addWidget(self.edit)

        row_btn = QHBoxLayout()
        self.btn_paste_current = QPushButton("この文字列をペースト")
        self.btn_copy_current = QPushButton("この文字列をクリップボードへ")
        row_btn.addWidget(self.btn_paste_current)
        row_btn.addWidget(self.btn_copy_current)
        if continuous_mode:
            self.btn_next_scan = QPushButton("続けてスキャン  ▶")
            self.btn_next_scan.setStyleSheet(
                "QPushButton { background-color: #1a7fd4; color: white; font-weight: bold; padding: 4px 12px; }"
                "QPushButton:hover { background-color: #1566b0; }"
            )
            row_btn.addWidget(self.btn_next_scan)
        row_btn.addStretch(1)
        lay.addLayout(row_btn)

        body = QHBoxLayout()

        left = QVBoxLayout()
        left.addWidget(QLabel("履歴から選択"))
        self.history_list = QListWidget()
        left.addWidget(self.history_list)
        self.btn_paste_history = QPushButton("選択した履歴をペースト")
        left.addWidget(self.btn_paste_history)
        body.addLayout(left, 1)

        right = QVBoxLayout()
        right.addWidget(QLabel("定型文字列から選択"))
        self.master_list = QListWidget()
        right.addWidget(self.master_list)
        self.btn_paste_master = QPushButton("選択した定型文字列をペースト")
        self.btn_edit_master = QPushButton("定型文字列を編集")
        right.addWidget(self.btn_paste_master)
        right.addWidget(self.btn_edit_master)
        body.addLayout(right, 1)

        lay.addLayout(body)

        bottom = QHBoxLayout()
        close_label = "スキャン終了" if continuous_mode else "閉じる"
        self.btn_cancel = QPushButton(close_label)
        bottom.addStretch(1)
        bottom.addWidget(self.btn_cancel)
        lay.addLayout(bottom)

        self.btn_paste_current.clicked.connect(self.paste_current)
        self.btn_copy_current.clicked.connect(self.copy_current)
        if continuous_mode:
            self.btn_next_scan.clicked.connect(self.next_scan)
        self.btn_paste_history.clicked.connect(self.paste_history)
        self.btn_paste_master.clicked.connect(self.paste_master)
        self.btn_edit_master.clicked.connect(self.edit_master)
        self.btn_cancel.clicked.connect(self.reject)

        self.reload_lists()

    def reload_lists(self):
        self.history_list.clear()
        for row in self.main_window.history.items:
            self.history_list.addItem(QListWidgetItem(row.get("text", "")))
        self.master_list.clear()
        for text in load_master_terms():
            self.master_list.addItem(QListWidgetItem(text))

    def maybe_learn(self, current_text):
        old = (self.raw_ocr or "").strip()
        new = (current_text or "").strip()
        if old and new and old != new:
            learn_replace_rule(old, new)

    def next_scan(self):
        """履歴に記録して次のスキャンへ進む（ペーストしない）。"""
        text = self.edit.toPlainText().strip()
        if not text:
            return
        self.maybe_learn(text)
        self.main_window.history.add(text, source="ocr")
        self.wants_next_scan = True
        self.accept()

    def _do_paste(self, text: str, source: str, stay: bool = False):
        """
        hide → pyautogui.hotkey("ctrl","v") → 必要ならダイアログへ戻る
        stay=True  : ダイアログへ戻る（履歴・定型文字列）
        stay=False : ダイアログを閉じてメインウィンドウへ（OCR結果）
        """
        text = (text or "").strip()
        if not text:
            return
        self.main_window.history.add(text, source=source)
        if stay:
            # exec() モーダルを抜けてからペーストし、完了後にダイアログを再オープン
            # （hide() のまま exec() が続くと Windows がダイアログをフォーカス保持と
            #   認識し、pyautogui のキーがそこに届いてしまうため）
            self.hide()
            self.accept()
            reopen = lambda: self.main_window.show_paste_dialog("")
            QTimer.singleShot(120, lambda: self.main_window.paste_text_then(text, reopen))
        else:
            # ダイアログを閉じてペースト（V3.0.0 と同じ2段階方式）
            self.hide()
            self.accept()
            QTimer.singleShot(120, lambda: self.main_window.paste_text(text))

    def paste_current(self):
        text = self.edit.toPlainText().strip()
        self.maybe_learn(text)
        self._do_paste(text, source="ocr", stay=False)

    def copy_current(self):
        text = self.edit.toPlainText().strip()
        if not text:
            return
        self.maybe_learn(text)
        pyperclip.copy(text)
        self.main_window.history.add(text, source="ocr")
        QMessageBox.information(self, "コピー", "クリップボードへコピーしました。")

    def paste_history(self):
        item = self.history_list.currentItem()
        if item:
            self._do_paste(item.text(), source="history", stay=True)

    def paste_master(self):
        item = self.master_list.currentItem()
        if item:
            self._do_paste(item.text(), source="master", stay=True)

    def edit_master(self):
        dlg = MasterTermsDialog(self)
        dlg.exec()
        self.reload_lists()


# ---------- Main window ----------
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.settings = Settings()
        self.history = HistoryStore(self.settings)
        self.overlay = CaptureOverlay(self.settings)
        self.overlay.captured.connect(self.on_ocr_captured)
        self.overlay.canceled.connect(self.on_capture_canceled)
        self.continuous_mode = False
        self._scanning = False

        self.setWindowTitle("画面OCR")
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.Tool)
        self.resize(360, 330)
        self.build_ui()
        self.setup_tray()
        self.apply_startup_setting(from_ui=False)
        self.restore_position()
        self.bootstrap_tesseract()
    def build_ui(self):
        lay = QVBoxLayout(self)

        title = QLabel("画面OCR")
        title.setStyleSheet("font-weight: bold; font-size: 16px;")
        lay.addWidget(title)

        self.btn_scan = QPushButton("画面スキャン")
        self.btn_cont = QPushButton("連続画面スキャン")
        self.btn_stop = QPushButton("スキャン中止")
        self.btn_history = QPushButton("履歴表示・編集")
        self.btn_paste = QPushButton("ペースト")
        self.btn_master = QPushButton("定型文字列編集")

        lay.addWidget(self.btn_scan)
        lay.addWidget(self.btn_cont)
        lay.addWidget(self.btn_stop)
        lay.addWidget(self.btn_history)
        lay.addWidget(self.btn_paste)
        lay.addWidget(self.btn_master)

        self.btn_stop.setEnabled(False)

        lay.addWidget(QLabel("補正設定"))
        self.chk_trim = QCheckBox("OCR文字列の空白を詰める")
        self.chk_chiseki = QCheckBox("登記事項の地積表記を補正")
        self.chk_chome = QCheckBox("住所の丁目は漢数字に置き換える")
        self.chk_fullwidth = QCheckBox("英数字は全角にする")
        self.chk_keep_half = QCheckBox("地積は半角のままにする")
        self.chk_startup = QCheckBox("Windows起動時に自動起動する")

        for w in [self.chk_trim, self.chk_chiseki, self.chk_chome, self.chk_fullwidth, self.chk_keep_half, self.chk_startup]:
            lay.addWidget(w)

        row_hist = QHBoxLayout()
        row_hist.addWidget(QLabel("履歴保存件数"))
        self.spin_hist = QSpinBox()
        self.spin_hist.setRange(1, 1000)
        row_hist.addWidget(self.spin_hist)
        lay.addLayout(row_hist)

        lay.addStretch(1)

        self.chk_trim.setChecked(self.settings["trim_spaces"])
        self.chk_chiseki.setChecked(self.settings["normalize_chiseki"])
        self.chk_chome.setChecked(self.settings["convert_chome"])
        self.chk_fullwidth.setChecked(self.settings["fullwidth_alnum"])
        self.chk_keep_half.setChecked(self.settings["keep_chiseki_halfwidth"])
        self.chk_startup.setChecked(self.settings["startup"])
        self.spin_hist.setValue(int(self.settings["history_limit"]))

        self.btn_scan.clicked.connect(self.start_single_scan)
        self.btn_cont.clicked.connect(self.start_continuous_scan)
        self.btn_stop.clicked.connect(self.stop_continuous_scan)
        self.btn_history.clicked.connect(self.open_history)
        self.btn_paste.clicked.connect(self.open_paste_dialog)
        self.btn_master.clicked.connect(self.open_master_dialog)

        self.chk_trim.toggled.connect(self.save_settings_from_ui)
        self.chk_chiseki.toggled.connect(self.save_settings_from_ui)
        self.chk_chome.toggled.connect(self.save_settings_from_ui)
        self.chk_fullwidth.toggled.connect(self.save_settings_from_ui)
        self.chk_keep_half.toggled.connect(self.save_settings_from_ui)
        self.chk_startup.toggled.connect(lambda _: self.apply_startup_setting(from_ui=True))
        self.spin_hist.valueChanged.connect(self.save_settings_from_ui)

    def setup_tray(self):
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.make_icon())
        self.tray.setToolTip("画面OCR")
        menu = QMenu()
        act_show = QAction("表示", self)
        act_scan = QAction("画面スキャン", self)
        act_quit = QAction("終了", self)
        act_show.triggered.connect(self.show_normal)
        act_scan.triggered.connect(self.start_single_scan)
        act_quit.triggered.connect(QApplication.instance().quit)
        menu.addAction(act_show)
        menu.addAction(act_scan)
        menu.addSeparator()
        menu.addAction(act_quit)
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(lambda reason: self.show_normal() if reason == QSystemTrayIcon.Trigger else None)
        self.tray.show()

    def make_icon(self):
        pm = QPixmap(QSize(32, 32))
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        p.setBrush(QColor(60, 120, 220))
        p.setPen(Qt.NoPen)
        p.drawRoundedRect(2, 2, 28, 28, 6, 6)
        p.setPen(QPen(Qt.white, 2))
        p.drawText(pm.rect(), Qt.AlignCenter, "OCR")
        p.end()
        return QIcon(pm)

    def restore_position(self):
        x = self.settings["window_x"]
        y = self.settings["window_y"]
        if x is not None and y is not None:
            self.move(int(x), int(y))
        else:
            screen = QGuiApplication.primaryScreen().availableGeometry()
            self.move(screen.right() - self.width() - 20, screen.top() + 20)

    def bootstrap_tesseract(self):
        exe_path, _ = configure_tesseract()
        if not exe_path:
            QMessageBox.warning(
                self,
                "Tesseract未検出",
                "同梱されているTesseractが見つかりませんでした。\n"
                "配布パッケージが正しく展開されているか確認してください。",
            )

    def moveEvent(self, event):
        self.settings["window_x"] = self.x()
        self.settings["window_y"] = self.y()
        self.settings.save()
        super().moveEvent(event)

    def closeEvent(self, event):
        event.ignore()
        self.hide()
        self.tray.showMessage("画面OCR", "右下トレイに常駐しています。", QSystemTrayIcon.Information, 1500)

    def show_normal(self):
        self.show()
        self.raise_()
        self.activateWindow()

    def save_settings_from_ui(self):
        self.settings["trim_spaces"] = self.chk_trim.isChecked()
        self.settings["normalize_chiseki"] = self.chk_chiseki.isChecked()
        self.settings["convert_chome"] = self.chk_chome.isChecked()
        self.settings["fullwidth_alnum"] = self.chk_fullwidth.isChecked()
        self.settings["keep_chiseki_halfwidth"] = self.chk_keep_half.isChecked()
        self.settings["history_limit"] = int(self.spin_hist.value())
        self.settings.save()
        self.history.trim()
        self.history.save()

    def apply_startup_setting(self, from_ui=False):
        if from_ui:
            self.save_settings_from_ui()
            self.settings["startup"] = self.chk_startup.isChecked()
            self.settings.save()
        enabled = bool(self.settings["startup"])
        try:
            if enabled:
                self.create_startup_cmd()
            else:
                if STARTUP_CMD.exists():
                    STARTUP_CMD.unlink()
        except Exception as e:
            QMessageBox.warning(self, "スタートアップ", f"設定に失敗しました。\n{e}")

    def create_startup_cmd(self):
        launch_target = get_app_launch_target()
        if is_frozen():
            cmd = f'@echo off\r\nstart "" "{launch_target}"\r\n'
        else:
            python_exe = Path(sys.executable).resolve()
            cmd = f'@echo off\r\nstart "" "{python_exe}" "{launch_target}"\r\n'
        STARTUP_CMD.parent.mkdir(parents=True, exist_ok=True)
        STARTUP_CMD.write_text(cmd, encoding="utf-8")

    def start_single_scan(self):
        self.continuous_mode = False
        self._scanning = True
        self.btn_stop.setEnabled(False)
        self.hide()
        QTimer.singleShot(150, self.overlay.begin_capture)

    def start_continuous_scan(self):
        self.continuous_mode = True
        self._scanning = True
        self.btn_stop.setEnabled(True)
        self.hide()
        QTimer.singleShot(150, lambda: self.overlay.begin_capture(continuous=True))

    def stop_continuous_scan(self):
        self.continuous_mode = False
        self._scanning = False
        self.btn_stop.setEnabled(False)
        self.show_normal()

    def on_capture_canceled(self):
        if self.continuous_mode and self._scanning:
            self.continuous_mode = False
            self._scanning = False
            self.btn_stop.setEnabled(False)
            self.show_normal()
            self.tray.showMessage("画面OCR", "連続スキャンを終了しました。", QSystemTrayIcon.Information, 2000)
        else:
            self.show_normal()

    def on_ocr_captured(self, raw_text: str):
        text, original, _chiseki = normalize_text(raw_text, self.settings)
        if not text.strip():
            self.show_normal()
            QMessageBox.information(self, "画面OCR", "文字を認識できませんでした。")
            if self.continuous_mode and self._scanning:
                self.hide()
                QTimer.singleShot(150, lambda: self.overlay.begin_capture(continuous=True))
            return
        self.show_paste_dialog(text, original)

    def show_paste_dialog(self, current_text: str, raw_ocr: str = ""):
        is_cont = self.continuous_mode and self._scanning
        dlg = PasteDialog(self, current_text=current_text, raw_ocr=raw_ocr, continuous_mode=is_cont)
        dlg.exec()
        if is_cont and dlg.wants_next_scan:
            self.hide()
            QTimer.singleShot(150, lambda: self.overlay.begin_capture(continuous=True))
        else:
            if self.continuous_mode:
                self.continuous_mode = False
                self._scanning = False
                self.btn_stop.setEnabled(False)
            self.show_normal()

    def open_paste_dialog(self):
        self.show_paste_dialog("")

    def open_history(self):
        dlg = HistoryDialog(self.history, self)
        dlg.exec()

    def open_master_dialog(self):
        dlg = MasterTermsDialog(self)
        dlg.exec()

    def paste_text(self, text: str):
        """ペーストしてメインウィンドウを再表示する。"""
        try:
            self.hide()
            pyperclip.copy(text)
            def do_paste():
                pyautogui.hotkey("ctrl", "v")
                QTimer.singleShot(300, self.show_normal)
            QTimer.singleShot(120, do_paste)
        except Exception as e:
            QMessageBox.warning(self, "ペースト", f"ペーストに失敗しました。\n{e}")
            self.show_normal()

    def paste_text_then(self, text: str, callback):
        """ペースト後に callback を呼ぶ（ペーストダイアログ再表示用）。"""
        try:
            self.hide()
            pyperclip.copy(text)
            def do_paste():
                pyautogui.hotkey("ctrl", "v")
                QTimer.singleShot(300, callback)
            QTimer.singleShot(120, do_paste)
        except Exception as e:
            QMessageBox.warning(self, "ペースト", f"ペーストに失敗しました。\n{e}")
            self.show_normal()


def main():
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()