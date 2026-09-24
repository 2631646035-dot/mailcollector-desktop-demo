"""
MailCollector - Outlook 邮件附件收集器
多邻国风格 | PySide6 | 方案记忆 | 多条件筛选 | 定时执行
"""

import sys, os, json, re, zipfile, shutil, uuid
from datetime import datetime, timezone, timedelta
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QStackedWidget, QPushButton, QLabel, QLineEdit, QTextEdit, QComboBox,
    QCheckBox, QFileDialog, QProgressBar, QScrollArea,
    QFrame, QSizePolicy, QMessageBox, QListWidget, QListWidgetItem,
    QDateEdit, QRadioButton, QButtonGroup, QDialog,
    QDialogButtonBox, QGraphicsDropShadowEffect, QTimeEdit, QSpinBox
)
from PySide6.QtCore import Qt, QDate, QThread, Signal, QSize, QTimer, QTime
from PySide6.QtGui import QFont, QColor

try:
    import win32com.client
except Exception:
    win32com = None


# ─── 配置 ───
APP_DIR = Path(os.path.expandvars(r"%APPDATA%")) / "MailCollector"
APP_DIR.mkdir(parents=True, exist_ok=True)
PLANS_FILE = APP_DIR / "plans.json"
DEFAULT_OUTPUT = Path.home() / "Desktop" / "MailCollector_Output"


# ─── 多邻国主题色 ───
C_PRIMARY = "#58CC02"
C_PRIMARY_DARK = "#46A302"
C_BG = "#FFFFFF"
C_CARD = "#F7F7F7"
C_TEXT = "#3C3C3C"
C_TEXT_SECONDARY = "#777777"
C_BORDER = "#E5E5E5"
C_ACCENT_BLUE = "#1CB0F6"
C_ACCENT_ORANGE = "#FF9600"
C_ACCENT_RED = "#FF4B4B"
C_WARNING = "#FFC800"


# ─── 全局样式 ───
STYLE = f"""
QMainWindow, QWidget {{
    background-color: {C_BG};
    color: {C_TEXT};
    font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
    font-size: 14px;
}}

QLabel {{
    color: {C_TEXT};
    background: transparent;
    border: none;
}}

QLineEdit, QTextEdit, QComboBox, QDateEdit, QTimeEdit, QSpinBox {{
    border: 2px solid {C_BORDER};
    border-radius: 14px;
    padding: 9px 13px;
    font-size: 14px;
    background: white;
    color: {C_TEXT};
    selection-background-color: {C_PRIMARY};
}}

QLineEdit:focus, QTextEdit:focus, QComboBox:focus, QDateEdit:focus, QTimeEdit:focus, QSpinBox:focus {{
    border-color: {C_PRIMARY};
}}

QLineEdit::placeholder {{
    color: #B0B0B0;
}}

QPushButton {{
    border: none;
    border-radius: 14px;
    padding: 10px 20px;
    font-size: 14px;
    font-weight: 700;
    background: {C_BORDER};
    color: {C_TEXT};
}}

QPushButton:hover {{
    background: #EFEFEF;
}}

QScrollArea {{
    border: none;
    background: transparent;
}}

QProgressBar {{
    border: none;
    border-radius: 8px;
    height: 10px;
    background: {C_BORDER};
    text-align: center;
}}

QProgressBar::chunk {{
    border-radius: 8px;
    background: {C_PRIMARY};
}}

QListWidget {{
    border: none;
    background: transparent;
    outline: none;
}}

QListWidget::item {{
    border: none;
    padding: 0px;
    margin: 6px 0px;
    background: transparent;
}}

QListWidget::item:hover {{
    background: transparent;
}}

QListWidget::item:selected {{
    background: transparent;
}}

QCheckBox {{
    font-size: 14px;
    spacing: 8px;
    background: transparent;
}}

QRadioButton {{
    font-size: 14px;
    spacing: 8px;
    background: transparent;
}}
"""


# ─── 通用 UI 组件 ───

def apply_shadow(widget, blur=22, y=6, color=QColor(0, 0, 0, 28)):
    shadow = QGraphicsDropShadowEffect(widget)
    shadow.setBlurRadius(blur)
    shadow.setOffset(0, y)
    shadow.setColor(color)
    widget.setGraphicsEffect(shadow)


class GreenButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setMinimumHeight(48)
        self.setStyleSheet(f"""
            QPushButton {{
                background: {C_PRIMARY};
                color: white;
                border: none;
                border-radius: 16px;
                padding: 12px 28px;
                font-size: 15px;
                font-weight: 800;
                border-bottom: 5px solid {C_PRIMARY_DARK};
            }}
            QPushButton:hover {{
                background: #61D80D;
            }}
            QPushButton:pressed {{
                border-bottom: 2px solid {C_PRIMARY_DARK};
                padding-top: 15px;
                padding-bottom: 9px;
            }}
        """)


class WhiteButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setMinimumHeight(48)
        self.setStyleSheet(f"""
            QPushButton {{
                background: white;
                color: {C_PRIMARY};
                border: 2px solid {C_BORDER};
                border-radius: 16px;
                padding: 12px 28px;
                font-size: 15px;
                font-weight: 800;
                border-bottom: 5px solid #D6D6D6;
            }}
            QPushButton:hover {{
                background: #F8FFF3;
                border-color: {C_PRIMARY};
            }}
            QPushButton:pressed {{
                border-bottom: 2px solid #D6D6D6;
                padding-top: 15px;
                padding-bottom: 9px;
            }}
        """)


class DangerButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setMinimumHeight(44)
        self.setStyleSheet(f"""
            QPushButton {{
                background: white;
                color: {C_ACCENT_RED};
                border: 2px solid {C_BORDER};
                border-radius: 16px;
                padding: 10px 22px;
                font-size: 13px;
                font-weight: 700;
                border-bottom: 4px solid #D6D6D6;
            }}
            QPushButton:hover {{
                background: #FFF0F0;
                border-color: {C_ACCENT_RED};
            }}
            QPushButton:pressed {{
                border-bottom: 2px solid #D6D6D6;
                padding-top: 12px;
                padding-bottom: 8px;
            }}
        """)


class Card(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("DuolingoCard")
        self.setStyleSheet(f"""
            QFrame#DuolingoCard {{
                background: {C_CARD};
                border-radius: 24px;
                border: none;
            }}
        """)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        apply_shadow(self, blur=20, y=5, color=QColor(0, 0, 0, 22))


class TitleLabel(QLabel):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setFont(QFont("Microsoft YaHei UI", 24, QFont.Bold))
        self.setStyleSheet(f"""
            QLabel {{
                color: {C_TEXT};
                background: transparent;
                border: none;
            }}
        """)


class SubtitleLabel(QLabel):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setFont(QFont("Microsoft YaHei UI", 12))
        self.setStyleSheet(f"""
            QLabel {{
                color: {C_TEXT_SECONDARY};
                background: transparent;
                border: none;
            }}
        """)


class SectionLabel(QLabel):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setFont(QFont("Microsoft YaHei UI", 15, QFont.Bold))
        self.setStyleSheet(f"""
            QLabel {{
                color: {C_TEXT};
                background: transparent;
                border: none;
            }}
        """)


class FieldLabel(QLabel):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setFont(QFont("Microsoft YaHei UI", 10, QFont.Bold))
        self.setStyleSheet(f"""
            QLabel {{
                color: {C_TEXT_SECONDARY};
                background: transparent;
                border: none;
                padding-left: 2px;
            }}
        """)


class StepBadge(QLabel):
    def __init__(self, text, color=C_PRIMARY, parent=None):
        super().__init__(text, parent)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setFont(QFont("Microsoft YaHei UI", 10, QFont.Bold))
        self.setStyleSheet(f"""
            QLabel {{
                background: white;
                color: {color};
                border: 2px solid #E5E5E5;
                border-radius: 14px;
                padding: 7px 12px;
            }}
        """)


class SmallActionButton(QPushButton):
    def __init__(self, text, color, parent=None):
        super().__init__(text, parent)
        self.setFixedSize(58, 32)
        self.setFont(QFont("Microsoft YaHei UI", 10, QFont.Bold))
        self.setStyleSheet(f"""
            QPushButton {{
                background: white;
                color: {color};
                border: 1px solid {color};
                border-radius: 10px;
                padding: 0px;
            }}
            QPushButton:hover {{
                background: {color};
                color: white;
            }}
        """)


class SegmentedControl(QWidget):
    value_changed = Signal(str)

    def __init__(self, items, default_value=None, parent=None):
        super().__init__(parent)

        self.group = QButtonGroup(self)
        self.group.setExclusive(True)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        self.buttons = []

        for value, text in items:
            btn = QPushButton(text)
            btn.setCheckable(True)
            btn.setProperty("value", value)
            btn.setMinimumHeight(42)
            btn.setStyleSheet(f"""
                QPushButton {{
                    background: white;
                    color: {C_TEXT_SECONDARY};
                    border: 2px solid {C_BORDER};
                    border-radius: 14px;
                    padding: 9px 16px;
                    font-size: 13px;
                    font-weight: 700;
                }}
                QPushButton:hover {{
                    border-color: {C_PRIMARY};
                    color: {C_PRIMARY};
                    background: #F8FFF3;
                }}
                QPushButton:checked {{
                    background: {C_PRIMARY};
                    color: white;
                    border-color: {C_PRIMARY};
                    border-bottom: 4px solid {C_PRIMARY_DARK};
                }}
            """)
            self.group.addButton(btn)
            layout.addWidget(btn)
            self.buttons.append(btn)
            btn.clicked.connect(lambda checked, v=value: self.value_changed.emit(v))

            if default_value is not None and value == default_value:
                btn.setChecked(True)

        if self.buttons and not any(b.isChecked() for b in self.buttons):
            self.buttons[0].setChecked(True)

    def value(self):
        checked = self.group.checkedButton()
        if checked:
            return checked.property("value")
        return None

    def set_value(self, value):
        for btn in self.buttons:
            if btn.property("value") == value:
                btn.setChecked(True)
                return
                return


class DateRangePicker(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        self.group = QButtonGroup(self)
        self.group.setExclusive(True)

        preset_row = QHBoxLayout()
        preset_row.setSpacing(8)

        self.buttons = []
        presets = [
            ("any", "不限"),
            ("today", "当天"),
            ("7d", "近 7 天"),
            ("30d", "近 30 天"),
            ("3m", "近 3 个月"),
            ("year", "今年"),
            ("custom", "自定义"),
        ]

        for value, text in presets:
            btn = QPushButton(text)
            btn.setCheckable(True)
            btn.setProperty("value", value)
            btn.setMinimumHeight(38)
            btn.setStyleSheet(f"""
                QPushButton {{
                    background: white;
                    color: {C_TEXT_SECONDARY};
                    border: 2px solid {C_BORDER};
                    border-radius: 19px;
                    padding: 7px 14px;
                    font-size: 13px;
                    font-weight: 700;
                }}
                QPushButton:hover {{
                    color: {C_PRIMARY};
                    border-color: {C_PRIMARY};
                    background: #F8FFF3;
                }}
                QPushButton:checked {{
                    background: {C_PRIMARY};
                    color: white;
                    border-color: {C_PRIMARY};
                }}
            """)
            self.group.addButton(btn)
            preset_row.addWidget(btn)
            self.buttons.append(btn)

        preset_row.addStretch()
        root.addLayout(preset_row)

        self.custom_box = QWidget()
        custom_row = QHBoxLayout(self.custom_box)
        custom_row.setContentsMargins(0, 0, 0, 0)
        custom_row.setSpacing(8)

        custom_row.addWidget(QLabel("从"))
        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDate(QDate.currentDate().addMonths(-3))
        self.date_from.setMinimumDate(QDate(2020, 1, 1))
        custom_row.addWidget(self.date_from)

        custom_row.addWidget(QLabel("到"))
        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate())
        self.date_to.setMinimumDate(QDate(2020, 1, 1))
        custom_row.addWidget(self.date_to)

        custom_row.addStretch()
        root.addWidget(self.custom_box)

        self.group.buttonClicked.connect(self._on_changed)
        self.set_value("3m")

    def _on_changed(self, *args):
        self.custom_box.setVisible(self.value() == "custom")

    def value(self):
        checked = self.group.checkedButton()
        if checked:
            return checked.property("value")
        return "3m"

    def set_value(self, value):
        for btn in self.buttons:
            if btn.property("value") == value:
                btn.setChecked(True)
                break
        self._on_changed()

    def set_range(self, date_from, date_to):
        if not date_from and not date_to:
            self.set_value("any")
            return

        self.set_value("custom")

        if date_from:
            self.date_from.setDate(QDate(date_from.year, date_from.month, date_from.day))

        if date_to:
            self.date_to.setDate(QDate(date_to.year, date_to.month, date_to.day))

    def get_range(self):
        mode = self.value()
        today = QDate.currentDate()

        if mode == "any":
            return None, None

        if mode == "today":
            start = today
            end = today
        elif mode == "7d":
            start = today.addDays(-7)
            end = today
        elif mode == "30d":
            start = today.addDays(-30)
            end = today
        elif mode == "3m":
            start = today.addMonths(-3)
            end = today
        elif mode == "year":
            start = QDate(today.year(), 1, 1)
            end = today
        else:
            start = self.date_from.date()
            end = self.date_to.date()
            if start > end:
                start, end = end, start

        date_from = datetime(start.year(), start.month(), start.day(), 0, 0, 0, tzinfo=timezone.utc)
        date_to = datetime(end.year(), end.month(), end.day(), 23, 59, 59, tzinfo=timezone.utc)

        return date_from, date_to


# ─── 工具函数 ───

def new_plan_id():
    return str(uuid.uuid4())


def action_to_text(action):
    return {
        "attachment": "下载附件",
        "email": "导出邮件",
        "both": "附件和邮件都要"
    }.get(action, action or "未知")


def title_mode_to_text(mode):
    return {
        "contains": "标题包含指定文字",
        "startswith": "标题以指定文字开头"
    }.get(mode, mode or "未设置")


def body_mode_to_text(mode):
    return {
        "any": "正文包含任一关键词",
        "all": "正文包含全部关键词"
    }.get(mode, mode or "未设置")


def date_mode_to_text(mode):
    return {
        "any": "不限",
        "today": "当天",
        "7d": "近 7 天",
        "30d": "近 30 天",
        "3m": "近 3 个月",
        "year": "今年",
        "custom": "自定义"
    }.get(mode, mode or "未设置")


def unique_path(path: Path):
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    i = 1

    while True:
        candidate = parent / f"{stem}_{i}{suffix}"
        if not candidate.exists():
            return candidate
        i += 1


def refresh_dynamic_dates(conditions):
    """
    如果保存的是“当天 / 近7天”等动态日期，每次执行前重新计算。
    """
    c = dict(conditions)
    mode = c.get("date_mode", "custom")

    today = QDate.currentDate()

    if mode == "any":
        c["date_from"] = None
        c["date_to"] = None
        return c

    if mode == "today":
        start = today
        end = today
    elif mode == "7d":
        start = today.addDays(-7)
        end = today
    elif mode == "30d":
        start = today.addDays(-30)
        end = today
    elif mode == "3m":
        start = today.addMonths(-3)
        end = today
    elif mode == "year":
        start = QDate(today.year(), 1, 1)
        end = today
    else:
        return c

    c["date_from"] = datetime(start.year(), start.month(), start.day(), 0, 0, 0, tzinfo=timezone.utc)
    c["date_to"] = datetime(end.year(), end.month(), end.day(), 23, 59, 59, tzinfo=timezone.utc)

    return c


def archive_current_files(output_dir, archive_mode):
    """
    定时执行前，把输出根目录中已有文件移动到“往期”。
    archive_mode:
      by_date: 每次创建日期文件夹
      common: 统一放到“往期”
    """
    root = Path(output_dir)
    root.mkdir(parents=True, exist_ok=True)

    archive_root = root / "往期"
    archive_root.mkdir(exist_ok=True)

    if archive_mode == "by_date":
        target = archive_root / datetime.now().strftime("%Y-%m-%d_%H%M%S")
        target.mkdir(exist_ok=True)
    else:
        target = archive_root

    for child in list(root.iterdir()):
        if child.name == "往期":
            continue

        dest = unique_path(target / child.name)

        try:
            shutil.move(str(child), str(dest))
        except Exception:
            pass


# ─── Outlook 邮件搜索与处理线程 ───

class OutlookWorker(QThread):
    progress = Signal(str)
    finished = Signal(dict)
    error = Signal(str)

    def __init__(self, conditions, output_dir, action):
        super().__init__()
        self.conditions = conditions
        self.output_dir = output_dir
        self.action = action

    def run(self):
        try:
            if win32com is None:
                raise RuntimeError("未检测到 pywin32。请先安装：pip install pywin32")

            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)

            self.progress.emit("🔍 正在搜索邮件...")

            messages = inbox.Items
            messages.Sort("[ReceivedTime]", True)

            matched = 0
            saved = 0
            skipped_duplicate = 0
            errors = []

            total = messages.Count
            self.progress.emit(f"📬 扫描 {total} 封邮件...")

            debug_count = 0
            all_mails = []

            for i in range(1, total + 1):
                try:
                    mail = messages.Item(i)
                except Exception:
                    continue

                try:
                    subject_preview = (mail.Subject or "")[:80]
                    sender_preview = mail.SenderEmailAddress or ""
                    date_preview = str(mail.ReceivedTime)[:19]
                    attachments_preview = [att.FileName[:80] for att in mail.Attachments]
                except Exception:
                    continue

                debug_count += 1
                ok, reason = self._match(mail)

                all_mails.append({
                    "subject": subject_preview,
                    "sender": sender_preview,
                    "date": date_preview,
                    "attachments": attachments_preview,
                    "matched": ok,
                    "reason": reason
                })

                if debug_count <= 30 and not ok:
                    self.progress.emit(f"✗ #{debug_count} {reason}: {subject_preview[:50]}")

                if not ok:
                    continue

                matched += 1
                self.progress.emit(f"✅ 匹配: {subject_preview[:50]}")

                if self.action in ("attachment", "both"):
                    exts = self.conditions.get("attachment_exts", [])
                    auto_unzip = self.conditions.get("auto_unzip", True)
                    zip_exts = {".zip", ".rar", ".7z"}

                    for att in mail.Attachments:
                        fname = att.FileName
                        _, file_ext = os.path.splitext(fname.lower())

                        if exts and file_ext not in [e.lower() for e in exts]:
                            continue

                        if self.conditions.get("_dedupe") and self._file_exists_anywhere(fname):
                            skipped_duplicate += 1
                            self.progress.emit(f"  ↩ 跳过重复附件: {fname}")
                            continue

                        save_path = os.path.join(self.output_dir, fname)
                        base, e = os.path.splitext(fname)
                        c = 1

                        while os.path.exists(save_path):
                            save_path = os.path.join(self.output_dir, f"{base}_{c}{e}")
                            c += 1

                        try:
                            att.SaveAsFile(save_path)
                            saved += 1
                            self.progress.emit(f"  📎 保存: {os.path.basename(save_path)}")

                            if auto_unzip and file_ext in zip_exts:
                                try:
                                    unzip_dir = os.path.join(self.output_dir, base)
                                    os.makedirs(unzip_dir, exist_ok=True)

                                    if file_ext == ".zip":
                                        with zipfile.ZipFile(save_path, "r") as zf:
                                            zf.extractall(unzip_dir)
                                    elif file_ext == ".rar":
                                        OutlookWorker._extract_rar(save_path, unzip_dir)
                                    elif file_ext == ".7z":
                                        OutlookWorker._extract_7z(save_path, unzip_dir)

                                    self.progress.emit(f"  📂 已解压到: {os.path.basename(unzip_dir)}")

                                except Exception as ze:
                                    self.progress.emit(f"  ⚠ 解压失败: {ze}")

                        except Exception as e:
                            errors.append(f"{fname}: {e}")

                if self.action in ("email", "both"):
                    fname = self._sanitize_filename(mail.Subject or "未命名邮件") + ".msg"

                    if self.conditions.get("_dedupe") and self._file_exists_anywhere(fname):
                        skipped_duplicate += 1
                        self.progress.emit(f"  ↩ 跳过重复邮件: {fname}")
                        continue

                    save_path = os.path.join(self.output_dir, fname)
                    c = 1

                    while os.path.exists(save_path):
                        fname = self._sanitize_filename(mail.Subject or "未命名邮件") + f"_{c}.msg"
                        save_path = os.path.join(self.output_dir, fname)
                        c += 1

                    try:
                        mail.SaveAs(save_path)
                        saved += 1
                        self.progress.emit(f"  ✉ 导出邮件: {fname}")
                    except Exception as e:
                        errors.append(f"邮件 {fname}: {e}")

                if matched >= 500:
                    self.progress.emit("⚠ 已达 500 封上限，停止扫描")
                    break

            fail_reasons = {}
            for m in all_mails:
                if not m["matched"]:
                    r = m["reason"]
                    fail_reasons[r] = fail_reasons.get(r, 0) + 1

            result = {
                "scanned": total,
                "matched": matched,
                "saved": saved,
                "skipped_duplicate": skipped_duplicate,
                "errors": errors,
                "output": self.output_dir,
                "conditions": self.conditions,
                "all_mails": all_mails,
                "fail_reasons": fail_reasons
            }

            self.finished.emit(result)

        except Exception as e:
            self.error.emit(str(e))

    def _file_exists_anywhere(self, filename):
        root = self.conditions.get("_dedupe_root")
        if not root:
            return False

        try:
            root_path = Path(root)
            if not root_path.exists():
                return False

            for p in root_path.rglob(filename):
                if p.is_file():
                    return True
        except Exception:
            return False

        return False

    def _normalize_dt(self, dt):
        try:
            if isinstance(dt, datetime):
                if dt.tzinfo is None:
                    return dt.replace(tzinfo=timezone.utc)
                return dt.astimezone(timezone.utc)
        except Exception:
            pass
        return dt

    def _match(self, mail):
        c = self.conditions

        try:
            subject = mail.Subject or ""
            body = mail.Body or ""
            sender = mail.SenderEmailAddress or ""
            recipients = ""

            try:
                recips = mail.Recipients
                recipients = ", ".join(str(r.Address) for r in recips if r.Address)
            except Exception:
                pass

            received = self._normalize_dt(mail.ReceivedTime)

        except Exception:
            return False, "parse_error"

        try:
            if c.get("date_from") and received < c["date_from"]:
                return False, "date_before_range"

            if c.get("date_to") and received > c["date_to"]:
                return False, "date_after_range"
        except Exception:
            return False, "date_compare_error"

        title_text = c.get("title_text", "").strip()
        if title_text:
            if c.get("title_mode") == "startswith":
                if not subject.lower().startswith(title_text.lower()):
                    return False, "title_startswith"
            else:
                if title_text.lower() not in subject.lower():
                    return False, "title_contains"

        body_text = c.get("body_text", "").strip()
        if body_text:
            keywords = [k.strip() for k in body_text.split(",") if k.strip()]
            if c.get("body_mode") == "any":
                if not any(k.lower() in body.lower() for k in keywords):
                    return False, "body_keyword"
            else:
                if not all(k.lower() in body.lower() for k in keywords):
                    return False, "body_keyword"

        sender_text = c.get("sender_text", "").strip()
        if sender_text:
            if sender_text.lower() not in sender.lower():
                return False, "sender"

        recip_text = c.get("recipient_text", "").strip()
        if recip_text:
            if recip_text.lower() not in recipients.lower():
                return False, "recipient"

        if c.get("has_attachments"):
            try:
                if mail.Attachments.Count == 0:
                    return False, "no_attachments"
            except Exception:
                return False, "attachments_error"

        return True, ""

    def _sanitize_filename(self, s):
        return re.sub(r'[\\/*?:"<>|]', "_", s)[:100]

    @staticmethod
    def _extract_rar(rar_path, out_dir):
        import subprocess

        unrar_paths = [
            r"C:\Program Files\WinRAR\UnRAR.exe",
            r"C:\Program Files (x86)\WinRAR\UnRAR.exe",
        ]

        for unrar in unrar_paths:
            if os.path.exists(unrar):
                subprocess.run([unrar, "x", "-y", rar_path, out_dir], capture_output=True, check=True)
                return

        raise RuntimeError("未找到 WinRAR/UnRAR，请安装 WinRAR 或使用 .zip 格式")

    @staticmethod
    def _extract_7z(sz_path, out_dir):
        import subprocess

        sz_paths = [
            r"C:\Program Files\7-Zip\7z.exe",
            r"C:\Program Files (x86)\7-Zip\7z.exe",
        ]

        for sz in sz_paths:
            if os.path.exists(sz):
                subprocess.run([sz, "x", f"-o{out_dir}", "-y", sz_path], capture_output=True, check=True)
                return

        raise RuntimeError("未找到 7-Zip，请安装 7-Zip 或使用 .zip 格式")


# ─── IMAP Worker ───

def resolve_imap_config(config):
    """Return a runtime-only IMAP configuration without persisting secrets."""
    resolved = dict(config or {})
    resolved["password"] = resolved.get("password") or os.getenv("MAILCOLLECTOR_IMAP_PASSWORD", "")
    return resolved

class IMAPWorker(QThread):
    progress = Signal(str)
    finished = Signal(dict)
    error = Signal(str)

    def __init__(self, conditions, output_dir, action, imap_config):
        super().__init__()
        self.conditions = conditions
        self.output_dir = output_dir
        self.action = action
        self.imap_config = resolve_imap_config(imap_config)

    def run(self):
        import imaplib
        import email
        from email.header import decode_header

        try:
            self.progress.emit("🔍 正在连接 IMAP 服务器...")

            server = self.imap_config["server"]
            port = self.imap_config["port"]
            account = self.imap_config["account"]
            password = self.imap_config["password"]

            conn = imaplib.IMAP4_SSL(server, port)
            conn.login(account, password)
            conn.select("INBOX")

            self.progress.emit("🔍 正在搜索邮件...")
            self.progress.emit(f"[条件] {self.conditions}")

            _, data = conn.search(None, "ALL")
            mail_ids = data[0].split()
            total = len(mail_ids)

            self.progress.emit(f"📬 扫描邮件: {total} 封")

            matched = 0
            saved = 0
            errors = []
            all_mails = []

            for i, mid in enumerate(mail_ids):
                try:
                    _, msg_data = conn.fetch(mid, "(RFC822)")
                    raw_email = msg_data[0][1]
                    msg = email.message_from_bytes(raw_email)
                except:
                    continue

                subject = self._decode_header(msg.get("Subject", ""))
                sender = msg.get("From", "")
                date_str = msg.get("Date", "")
                body = self._get_body(msg)

                has_att = False
                for part in msg.walk():
                    if part.get_content_disposition() == "attachment":
                        has_att = True
                        break

                received = None
                if date_str:
                    try:
                        from email.utils import parsedate_to_datetime
                        received = parsedate_to_datetime(date_str)
                        if received.tzinfo is None:
                            received = received.replace(tzinfo=timezone.utc)
                    except:
                        pass

                ok, reason = self._match_imap(subject, body, sender, received, has_att)

                all_mails.append({
                    "subject": subject[:80],
                    "sender": sender,
                    "date": str(received)[:19] if received else "",
                    "matched": ok,
                    "reason": reason
                })

                if not ok:
                    if len(all_mails) <= 30 and not ok:
                        self.progress.emit(f"✗ #{len(all_mails)} {reason}: {subject[:50]}")
                    continue

                matched += 1
                self.progress.emit(f"✅ 匹配: {subject[:50]}")

                if self.action in ("attachment", "both"):
                    exts = self.conditions.get("attachment_exts", [])
                    auto_unzip = self.conditions.get("auto_unzip", True)
                    ZIP_EXTS = {'.zip', '.rar', '.7z'}

                    for part in msg.walk():
                        if part.get_content_disposition() != "attachment":
                            continue

                        fname = part.get_filename()
                        if not fname:
                            continue
                        fname = self._decode_header(fname)

                        _, file_ext = os.path.splitext(fname.lower())
                        if exts and file_ext not in [e.lower() for e in exts]:
                            continue

                        save_path = os.path.join(self.output_dir, fname)
                        base, e = os.path.splitext(fname)
                        c = 1
                        while os.path.exists(save_path):
                            save_path = os.path.join(self.output_dir, f"{base}_{c}{e}")
                            c += 1

                        try:
                            with open(save_path, "wb") as f:
                                f.write(part.get_payload(decode=True))
                            saved += 1
                            self.progress.emit(f"  📎 保存: {os.path.basename(save_path)}")

                            if auto_unzip and file_ext in ZIP_EXTS:
                                try:
                                    unzip_dir = os.path.join(self.output_dir, base)
                                    os.makedirs(unzip_dir, exist_ok=True)
                                    if file_ext == '.zip':
                                        with zipfile.ZipFile(save_path, 'r') as zf:
                                            zf.extractall(unzip_dir)
                                    elif file_ext == '.rar':
                                        OutlookWorker._extract_rar(None, save_path, unzip_dir)
                                    elif file_ext == '.7z':
                                        OutlookWorker._extract_7z(None, save_path, unzip_dir)
                                    self.progress.emit(f"  📂 已解压到: {os.path.basename(unzip_dir)}")
                                except Exception as ze:
                                    self.progress.emit(f"  ⚠ 解压失败: {ze}")
                        except Exception as ex:
                            errors.append(f"{fname}: {ex}")

            fail_reasons = {}
            for m in all_mails:
                if not m["matched"]:
                    r = m["reason"]
                    fail_reasons[r] = fail_reasons.get(r, 0) + 1

            result = {
                "scanned": total,
                "matched": matched,
                "saved": saved,
                "errors": errors,
                "output": self.output_dir,
                "conditions": self.conditions,
                "all_mails": all_mails,
                "fail_reasons": fail_reasons
            }

            conn.close()
            conn.logout()
            self.finished.emit(result)

        except Exception as e:
            self.error.emit(str(e))

    def _decode_header(self, h):
        from email.header import decode_header
        try:
            parts = decode_header(h)
            result = []
            for part, charset in parts:
                if isinstance(part, bytes):
                    result.append(part.decode(charset or "utf-8", errors="replace"))
                else:
                    result.append(str(part))
            return "".join(result)
        except:
            return str(h)

    def _get_body(self, msg):
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    try:
                        charset = part.get_content_charset() or "utf-8"
                        body += part.get_payload(decode=True).decode(charset, errors="replace")
                    except:
                        pass
        else:
            try:
                charset = msg.get_content_charset() or "utf-8"
                body = msg.get_payload(decode=True).decode(charset, errors="replace")
            except:
                body = str(msg.get_payload())
        return body

    def _match_imap(self, subject, body, sender, received, has_att):
        c = self.conditions

        if c.get("date_from") and received:
            if received < c["date_from"]:
                return (False, "date_before_range")
        if c.get("date_to") and received:
            if received > c["date_to"]:
                return (False, "date_after_range")

        title_text = c.get("title_text", "").strip()
        if title_text:
            if c.get("title_mode") == "startswith":
                if not subject.lower().startswith(title_text.lower()):
                    return (False, "title_startswith")
            else:
                if title_text.lower() not in subject.lower():
                    return (False, "title_contains")

        body_text = c.get("body_text", "").strip()
        if body_text:
            keywords = [k.strip() for k in body_text.split(",") if k.strip()]
            if c.get("body_mode") == "any":
                if not any(k.lower() in body.lower() for k in keywords):
                    return (False, "body_keyword")
            else:
                if not all(k.lower() in body.lower() for k in keywords):
                    return (False, "body_keyword")

        sender_text = c.get("sender_text", "").strip()
        if sender_text:
            if sender_text.lower() not in sender.lower():
                return (False, "sender")

        if c.get("has_attachments") and not has_att:
            return (False, "no_attachments")

        return (True, "")


# ─── 方案管理 ───

def load_plans():
    changed = False

    if PLANS_FILE.exists():
        with open(PLANS_FILE, "r", encoding="utf-8") as f:
            plans = json.load(f)

        for plan in plans:
            if not plan.get("id"):
                plan["id"] = new_plan_id()
                changed = True

            if "schedule" not in plan:
                plan["schedule"] = {
                    "enabled": False,
                    "time": "09:00",
                    "months": 1,
                    "end_date": "",
                    "archive_mode": "by_date",
                    "last_run_date": ""
                }
                changed = True

            c = plan.get("conditions", {})

            if "date_mode" not in c:
                c["date_mode"] = "custom"
                changed = True

            for key in ("date_from", "date_to"):
                val = c.get(key)
                if val:
                    try:
                        c[key] = datetime.fromisoformat(val)
                        if c[key].tzinfo is None:
                            c[key] = c[key].replace(tzinfo=timezone.utc)
                        else:
                            c[key] = c[key].astimezone(timezone.utc)
                    except Exception:
                        c[key] = None

        if changed:
            save_plans(plans)

        return plans

    return []


def save_plans(plans):
    def convert(obj):
        if isinstance(obj, datetime):
            return obj.isoformat()
        if isinstance(obj, dict):
            return {k: convert(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [convert(v) for v in obj]
        return obj

    safe_plans = convert(plans)
    for plan in safe_plans:
        imap_config = plan.get("imap_config")
        if isinstance(imap_config, dict):
            imap_config.pop("password", None)

    with open(PLANS_FILE, "w", encoding="utf-8") as f:
        json.dump(safe_plans, f, ensure_ascii=False, indent=2)


# ─── 多邻国风格弹窗 ───

class StyledDialog(QDialog):
    def __init__(self, title="", fixed_size=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)

        if fixed_size:
            self.setFixedSize(*fixed_size)

        self.setStyleSheet(f"""
            QDialog {{
                background: {C_BG};
            }}
            QLabel {{
                background: transparent;
                border: none;
            }}
            QRadioButton, QCheckBox {{
                font-size: 14px;
                font-weight: 700;
                color: {C_TEXT};
                background: transparent;
            }}
            QDialogButtonBox QPushButton {{
                background: white;
                color: {C_PRIMARY};
                border: 2px solid {C_BORDER};
                border-radius: 14px;
                padding: 8px 18px;
                font-size: 13px;
                font-weight: 800;
                min-width: 70px;
            }}
            QDialogButtonBox QPushButton:hover {{
                border-color: {C_PRIMARY};
                background: #F8FFF3;
            }}
        """)


class SavePlanDialog(StyledDialog):
    def __init__(self, default_name="", editing=False, parent=None):
        super().__init__("保存方案", fixed_size=(480, 250), parent=parent)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(26, 22, 26, 22)
        layout.setSpacing(14)

        head = QHBoxLayout()

        icon = QLabel("💾")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setFixedSize(48, 48)
        icon.setStyleSheet("""
            QLabel {
                background: #E5F8D8;
                border-radius: 24px;
                font-size: 24px;
            }
        """)
        head.addWidget(icon)

        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        title_box.addWidget(SectionLabel("保存这个方案"))
        title_box.addWidget(SubtitleLabel("可以覆盖原方案，也可以另存为一个新方案"))
        head.addLayout(title_box)
        head.addStretch()

        layout.addLayout(head)

        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("例如：法国一件代发")
        self.name_edit.setText(default_name)
        layout.addWidget(self.name_edit)

        self.overwrite_rb = QRadioButton("覆盖原方案")
        self.save_as_rb = QRadioButton("另存为新方案")

        self.mode_group = QButtonGroup(self)
        self.mode_group.addButton(self.overwrite_rb)
        self.mode_group.addButton(self.save_as_rb)

        if editing:
            self.overwrite_rb.setChecked(True)
        else:
            self.overwrite_rb.setEnabled(False)
            self.save_as_rb.setChecked(True)

        layout.addWidget(self.overwrite_rb)
        layout.addWidget(self.save_as_rb)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def result_data(self):
        mode = "overwrite" if self.overwrite_rb.isChecked() else "save_as"
        return mode, self.name_edit.text().strip()


class PlanDetailDialog(StyledDialog):
    def __init__(self, plan, parent=None):
        super().__init__("方案明细", fixed_size=(620, 560), parent=parent)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(26, 22, 26, 22)
        layout.setSpacing(14)

        head = QHBoxLayout()

        icon = QLabel("📋")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setFixedSize(50, 50)
        icon.setStyleSheet("""
            QLabel {
                background: #E8F5FF;
                border-radius: 25px;
                font-size: 24px;
            }
        """)
        head.addWidget(icon)

        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        title_box.addWidget(SectionLabel(plan.get("name", "未命名方案")))
        title_box.addWidget(SubtitleLabel("这个方案会按照下面的规则执行"))
        head.addLayout(title_box)
        head.addStretch()

        layout.addLayout(head)

        card = QFrame()
        card.setStyleSheet(f"""
            QFrame {{
                background: {C_CARD};
                border-radius: 20px;
                border: none;
            }}
        """)
        apply_shadow(card, blur=16, y=4, color=QColor(0, 0, 0, 18))

        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(18, 16, 18, 16)
        card_layout.setSpacing(10)

        text = QTextEdit()
        text.setReadOnly(True)
        text.setStyleSheet(f"""
            QTextEdit {{
                background: white;
                border: 2px solid {C_BORDER};
                border-radius: 16px;
                padding: 12px;
                font-size: 13px;
                color: {C_TEXT};
            }}
        """)

        c = plan.get("conditions", {})
        schedule = plan.get("schedule", {})

        detail = ""
        detail += "━━━ 基本信息 ━━━\n"
        detail += f"处理方式：{action_to_text(plan.get('action'))}\n"
        detail += f"保存位置：{plan.get('output_dir', '')}\n"
        detail += f"创建时间：{plan.get('created', '')}\n\n"

        detail += "━━━ 邮件筛选 ━━━\n"
        detail += f"标题规则：{title_mode_to_text(c.get('title_mode'))}\n"
        detail += f"标题文字：{c.get('title_text') or '未设置'}\n"
        detail += f"正文规则：{body_mode_to_text(c.get('body_mode'))}\n"
        detail += f"正文关键词：{c.get('body_text') or '未设置'}\n"
        detail += f"发件人包含：{c.get('sender_text') or '未设置'}\n"
        detail += f"收件人包含：{c.get('recipient_text') or '未设置'}\n"
        detail += f"日期范围：{date_mode_to_text(c.get('date_mode', 'custom'))}\n"
        detail += f"仅有附件：{'是' if c.get('has_attachments') else '否'}\n\n"

        detail += "━━━ 附件处理 ━━━\n"
        detail += f"附件类型：{', '.join(c.get('attachment_exts', [])) or '不限'}\n"
        detail += f"自动解压：{'是' if c.get('auto_unzip') else '否'}\n\n"

        detail += "━━━ 定时执行 ━━━\n"
        detail += f"是否启用：{'是' if schedule.get('enabled') else '否'}\n"
        detail += f"执行时间：{schedule.get('time', '09:00')}\n"
        detail += f"运行期限：{schedule.get('months', 1)} 个月\n"
        detail += f"结束日期：{schedule.get('end_date') or '未设置'}\n"
        detail += f"往期方式：{'每次按日期建文件夹' if schedule.get('archive_mode') == 'by_date' else '统一放入往期文件夹'}\n"
        detail += f"上次执行：{schedule.get('last_run_date') or '未执行'}\n"

        text.setText(detail)
        card_layout.addWidget(text)

        layout.addWidget(card)

        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        btns.rejected.connect(self.reject)
        btns.accepted.connect(self.accept)
        layout.addWidget(btns)


class ScheduleDialog(StyledDialog):
    def __init__(self, plan, parent=None):
        super().__init__("设置定时执行", fixed_size=(520, 390), parent=parent)

        self.plan = plan
        schedule = plan.get("schedule", {})

        layout = QVBoxLayout(self)
        layout.setContentsMargins(26, 22, 26, 22)
        layout.setSpacing(14)

        head = QHBoxLayout()

        icon = QLabel("⏰")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setFixedSize(50, 50)
        icon.setStyleSheet("""
            QLabel {
                background: #FFF4D6;
                border-radius: 25px;
                font-size: 24px;
            }
        """)
        head.addWidget(icon)

        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        title_box.addWidget(SectionLabel(f"定时执行：{plan.get('name', '')}"))
        title_box.addWidget(SubtitleLabel("程序保持打开时，会按指定时间自动执行"))
        head.addLayout(title_box)
        head.addStretch()

        layout.addLayout(head)

        card = QFrame()
        card.setStyleSheet(f"""
            QFrame {{
                background: {C_CARD};
                border-radius: 20px;
                border: none;
            }}
        """)
        apply_shadow(card, blur=16, y=4, color=QColor(0, 0, 0, 18))

        cl = QVBoxLayout(card)
        cl.setContentsMargins(18, 16, 18, 16)
        cl.setSpacing(12)

        self.enable_cb = QCheckBox("  ✅ 启用定时执行")
        self.enable_cb.setCursor(Qt.CursorShape.PointingHandCursor)

        # 默认打开“定时”弹窗时就是启用状态，用户如果不想启用可以再点一次取消
        self.enable_cb.setChecked(True)

        self.enable_cb.setMinimumHeight(58)
        self.enable_cb.setStyleSheet(f"""
            QCheckBox {{
                background: white;
                color: {C_TEXT};
                border: 2px solid {C_BORDER};
                border-radius: 18px;
                padding: 14px 18px;
                font-size: 15px;
                font-weight: 800;
            }}

            QCheckBox:hover {{
                border-color: {C_PRIMARY};
                background: #F8FFF3;
                color: {C_PRIMARY};
            }}

            QCheckBox:checked {{
                background: #E5F8D8;
                color: {C_PRIMARY_DARK};
                border: 2px solid {C_PRIMARY};
            }}

            QCheckBox::indicator {{
                width: 22px;
                height: 22px;
            }}
        """)

        cl.addWidget(self.enable_cb)

        tip_lbl = QLabel("点击上方绿色框即可开启或关闭定时执行")
        tip_lbl.setStyleSheet(f"""
            QLabel {{
                color: {C_TEXT_SECONDARY};
                font-size: 12px;
                padding-left: 4px;
            }}
        """)
        cl.addWidget(tip_lbl)

        row1 = QHBoxLayout()
        row1.addWidget(FieldLabel("每天执行时间"))

        self.time_edit = QTimeEdit()
        self.time_edit.setDisplayFormat("HH:mm")

        try:
            h, m = schedule.get("time", "09:00").split(":")
            self.time_edit.setTime(QTime(int(h), int(m)))
        except Exception:
            self.time_edit.setTime(QTime(9, 0))

        row1.addWidget(self.time_edit)
        row1.addStretch()
        cl.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(FieldLabel("运行多久"))

        self.month_spin = QSpinBox()
        self.month_spin.setMinimum(1)
        self.month_spin.setMaximum(6)
        self.month_spin.setValue(int(schedule.get("months", 1)))
        self.month_spin.setSuffix(" 个月")

        row2.addWidget(self.month_spin)
        row2.addStretch()
        cl.addLayout(row2)

        cl.addWidget(FieldLabel("往期文件处理方式"))

        self.archive_group = QButtonGroup(self)
        self.archive_by_date = QRadioButton("每次执行都创建一个日期文件夹")
        self.archive_common = QRadioButton("所有往期文件统一放在「往期」文件夹中")

        self.archive_group.addButton(self.archive_by_date)
        self.archive_group.addButton(self.archive_common)

        if schedule.get("archive_mode", "by_date") == "common":
            self.archive_common.setChecked(True)
        else:
            self.archive_by_date.setChecked(True)

        cl.addWidget(self.archive_by_date)
        cl.addWidget(self.archive_common)

        hint = QLabel("说明：定时任务执行前，会先把输出目录中的旧文件移入「往期」。新收集到的文件会留在输出目录根目录。")
        hint.setWordWrap(True)
        hint.setStyleSheet(f"color: {C_TEXT_SECONDARY}; font-size: 12px;")
        cl.addWidget(hint)

        layout.addWidget(card)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def result_data(self):
        months = self.month_spin.value()
        end_date = (datetime.now() + timedelta(days=months * 30)).strftime("%Y-%m-%d")

        return {
            "enabled": self.enable_cb.isChecked(),
            "time": self.time_edit.time().toString("HH:mm"),
            "months": months,
            "end_date": end_date,
            "archive_mode": "by_date" if self.archive_by_date.isChecked() else "common",
            "last_run_date": self.plan.get("schedule", {}).get("last_run_date", "")
        }


# ─── 页面1: 收集助手 ───

class CreateTaskPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main = main_window

        self.editing_plan_id = None
        self.editing_plan_name = ""

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(18)
        layout.setContentsMargins(40, 30, 40, 30)

        header = QHBoxLayout()
        header.setSpacing(14)

        mascot = QLabel("🦉")
        mascot.setAlignment(Qt.AlignmentFlag.AlignCenter)
        mascot.setFixedSize(58, 58)
        mascot.setStyleSheet("""
            QLabel {
                background: #E5F8D8;
                border-radius: 29px;
                font-size: 30px;
                border: none;
            }
        """)
        header.addWidget(mascot)

        title_box = QVBoxLayout()
        title_box.setSpacing(4)
        title_box.addWidget(TitleLabel("邮件收集助手"))
        subtitle = SubtitleLabel("从 Outlook 自动收集邮件、附件和压缩包")
        subtitle.setObjectName("subtitle")
        title_box.addWidget(subtitle)
        header.addLayout(title_box)
        header.addStretch()

        layout.addLayout(header)

        # 邮箱类型选择
        source_card = Card()
        scl = QVBoxLayout(source_card)
        scl.setContentsMargins(24, 16, 24, 16)
        scl.setSpacing(10)
        scl.addWidget(FieldLabel("邮箱来源"))
        source_row = QHBoxLayout()
        source_row.setSpacing(12)
        self.mail_source = SegmentedControl([
            ("outlook", "Outlook 客户端"),
            ("imap", "IMAP 邮箱"),
        ], default_value="outlook")
        self.mail_source.value_changed.connect(self._on_source_changed)
        source_row.addWidget(self.mail_source)
        source_row.addStretch()
        scl.addLayout(source_row)

        # IMAP 配置面板（默认隐藏）
        self.imap_panel = QWidget()
        self.imap_panel.setVisible(False)
        imap_layout = QVBoxLayout(self.imap_panel)
        imap_layout.setContentsMargins(0, 6, 0, 0)
        imap_layout.setSpacing(8)

        imap_row1 = QHBoxLayout()
        imap_row1.setSpacing(10)
        imap_left = QVBoxLayout()
        imap_left.setSpacing(4)
        imap_left.addWidget(FieldLabel("IMAP 服务器"))
        self.imap_server = QLineEdit()
        self.imap_server.setPlaceholderText("例如: imap.gmail.com")
        imap_left.addWidget(self.imap_server)
        imap_row1.addLayout(imap_left)
        imap_right = QVBoxLayout()
        imap_right.setSpacing(4)
        imap_right.addWidget(FieldLabel("端口"))
        self.imap_port = QLineEdit()
        self.imap_port.setPlaceholderText("993")
        self.imap_port.setText("993")
        self.imap_port.setFixedWidth(100)
        imap_right.addWidget(self.imap_port)
        imap_row1.addLayout(imap_right)
        imap_layout.addLayout(imap_row1)

        imap_row2 = QHBoxLayout()
        imap_row2.setSpacing(10)
        imap_acc = QVBoxLayout()
        imap_acc.setSpacing(4)
        imap_acc.addWidget(FieldLabel("邮箱账号"))
        self.imap_account = QLineEdit()
        self.imap_account.setPlaceholderText("demo.user@example.com")
        imap_acc.addWidget(self.imap_account)
        imap_row2.addLayout(imap_acc)
        imap_pwd = QVBoxLayout()
        imap_pwd.setSpacing(4)
        imap_pwd.addWidget(FieldLabel("密码/授权码"))
        self.imap_password = QLineEdit()
        self.imap_password.setEchoMode(QLineEdit.EchoMode.Password)
        self.imap_password.setPlaceholderText("请输入密码或应用授权码")
        imap_pwd.addWidget(self.imap_password)
        imap_row2.addLayout(imap_pwd)
        imap_layout.addLayout(imap_row2)

        scl.addWidget(self.imap_panel)

        layout.addWidget(source_card)

        steps = QHBoxLayout()
        steps.setSpacing(8)
        steps.addWidget(StepBadge("1 选邮件", C_PRIMARY))
        steps.addWidget(StepBadge("2 选文件", C_ACCENT_BLUE))
        steps.addWidget(StepBadge("3 自动收集", C_ACCENT_ORANGE))
        steps.addStretch()
        layout.addLayout(steps)

        card1 = Card()
        cl1 = QVBoxLayout(card1)
        cl1.setContentsMargins(24, 22, 24, 22)
        cl1.setSpacing(14)

        cl1.addWidget(SectionLabel("邮件筛选条件"))

        cl1.addWidget(FieldLabel("邮件标题"))
        title_mode_row = QHBoxLayout()
        title_mode_row.setSpacing(8)

        self.title_mode = SegmentedControl([
            ("contains", "包含文字"),
            ("startswith", "以文字开头"),
        ], default_value="contains")

        title_mode_row.addWidget(self.title_mode)

        self.title_text = QLineEdit()
        self.title_text.setPlaceholderText("例如：Invoice、账单、报价单")
        title_mode_row.addWidget(self.title_text, 1)

        cl1.addLayout(title_mode_row)

        cl1.addWidget(FieldLabel("邮件正文关键词"))
        body_row = QHBoxLayout()
        body_row.setSpacing(8)

        self.body_mode = SegmentedControl([
            ("any", "任一关键词"),
            ("all", "全部关键词"),
        ], default_value="any")

        body_row.addWidget(self.body_mode)

        self.body_text = QLineEdit()
        self.body_text.setPlaceholderText("多个关键词用英文逗号分隔，例如：invoice, payment, 账单")
        body_row.addWidget(self.body_text, 1)

        cl1.addLayout(body_row)

        row2 = QHBoxLayout()
        row2.setSpacing(12)

        left = QVBoxLayout()
        left.setSpacing(6)
        left.addWidget(FieldLabel("发件人包含"))
        self.sender_text = QLineEdit()
        self.sender_text.setPlaceholderText("例如：@supplier.example")
        left.addWidget(self.sender_text)
        row2.addLayout(left)

        right = QVBoxLayout()
        right.setSpacing(6)
        right.addWidget(FieldLabel("收件人包含"))
        self.recipient_text = QLineEdit()
        self.recipient_text.setPlaceholderText("例如：@buyer.example")
        right.addWidget(self.recipient_text)
        row2.addLayout(right)

        cl1.addLayout(row2)

        cl1.addWidget(FieldLabel("邮件日期范围"))
        self.date_picker = DateRangePicker()
        cl1.addWidget(self.date_picker)

        self.has_attachment = QCheckBox("仅筛选有附件的邮件")
        self.has_attachment.setChecked(True)
        self.has_attachment.setStyleSheet(f"""
            QCheckBox {{
                font-size: 14px;
                font-weight: 700;
                color: {C_TEXT};
                padding-top: 4px;
            }}
        """)
        cl1.addWidget(self.has_attachment)

        layout.addWidget(card1)

        card2 = Card()
        cl2 = QVBoxLayout(card2)
        cl2.setContentsMargins(24, 22, 24, 22)
        cl2.setSpacing(14)

        cl2.addWidget(SectionLabel("收集内容与保存方式"))

        cl2.addWidget(FieldLabel("你想收集什么"))
        self.action_group = QButtonGroup(self)
        act_row = QHBoxLayout()
        act_row.setSpacing(10)

        self.action_attachment = QRadioButton("下载附件")
        self.action_email = QRadioButton("导出邮件 .msg")
        self.action_both = QRadioButton("附件和邮件都要")
        self.action_both.setChecked(True)

        for rb in [self.action_attachment, self.action_email, self.action_both]:
            rb.setStyleSheet(f"""
                QRadioButton {{
                    font-size: 14px;
                    font-weight: 700;
                    color: {C_TEXT};
                    background: white;
                    border: 2px solid {C_BORDER};
                    border-radius: 16px;
                    padding: 10px 14px;
                }}
                QRadioButton:hover {{
                    border-color: {C_PRIMARY};
                    background: #F8FFF3;
                }}
            """)
            self.action_group.addButton(rb)
            act_row.addWidget(rb)

        act_row.addStretch()
        cl2.addLayout(act_row)

        cl2.addWidget(FieldLabel("附件类型"))
        ext_row = QHBoxLayout()
        ext_row.setSpacing(8)

        self.attachment_ext_presets = QCheckBox(".pdf")
        self.attachment_ext_presets.setChecked(True)
        ext_row.addWidget(self.attachment_ext_presets)

        self.attachment_ext_doc = QCheckBox(".doc/.docx")
        ext_row.addWidget(self.attachment_ext_doc)

        self.attachment_ext_xls = QCheckBox(".xls/.xlsx")
        ext_row.addWidget(self.attachment_ext_xls)

        self.attachment_ext_zip = QCheckBox(".zip/.rar/.7z")
        self.attachment_ext_zip.setChecked(True)
        ext_row.addWidget(self.attachment_ext_zip)

        self.attachment_ext_img = QCheckBox(".png/.jpg")
        ext_row.addWidget(self.attachment_ext_img)

        for cb in [
            self.attachment_ext_presets,
            self.attachment_ext_doc,
            self.attachment_ext_xls,
            self.attachment_ext_zip,
            self.attachment_ext_img
        ]:
            self._style_chip(cb)

        ext_row.addStretch()
        cl2.addLayout(ext_row)

        custom_row = QHBoxLayout()
        custom_row.setSpacing(8)
        custom_row.addWidget(FieldLabel("自定义后缀"))

        self.attachment_ext_custom = QLineEdit()
        self.attachment_ext_custom.setPlaceholderText(".msg, .txt，英文逗号分隔；留空表示不限")
        self.attachment_ext_custom.setFixedWidth(360)
        custom_row.addWidget(self.attachment_ext_custom)
        custom_row.addStretch()

        cl2.addLayout(custom_row)

        unzip_row = QHBoxLayout()
        self.auto_unzip = QCheckBox("自动解压压缩包 .zip / .rar / .7z")
        self.auto_unzip.setChecked(True)
        self.auto_unzip.setStyleSheet(f"""
            QCheckBox {{
                font-size: 14px;
                font-weight: 700;
                color: {C_TEXT};
            }}
        """)
        unzip_row.addWidget(self.auto_unzip)
        unzip_row.addStretch()
        cl2.addLayout(unzip_row)

        cl2.addWidget(FieldLabel("保存位置"))
        out_row = QHBoxLayout()
        out_row.setSpacing(8)

        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setText(str(DEFAULT_OUTPUT))
        out_row.addWidget(self.output_dir_edit)

        browse_btn = QPushButton("📁 选择")
        browse_btn.setMinimumHeight(42)
        browse_btn.setStyleSheet(f"""
            QPushButton {{
                background: white;
                color: {C_TEXT};
                border: 2px solid {C_BORDER};
                border-radius: 14px;
                padding: 9px 18px;
                font-size: 13px;
                font-weight: 800;
            }}
            QPushButton:hover {{
                border-color: {C_PRIMARY};
                color: {C_PRIMARY};
                background: #F8FFF3;
            }}
        """)
        browse_btn.clicked.connect(self._browse_output)
        out_row.addWidget(browse_btn)

        cl2.addLayout(out_row)

        layout.addWidget(card2)

        btn_row = QHBoxLayout()
        btn_row.addStretch()

        save_plan_btn = WhiteButton("💾 保存为方案")
        save_plan_btn.clicked.connect(self._save_plan)
        btn_row.addWidget(save_plan_btn)

        run_btn = GreenButton("🚀 立即执行")
        run_btn.clicked.connect(self._run)
        btn_row.addWidget(run_btn)

        layout.addLayout(btn_row)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setFixedHeight(10)
        layout.addWidget(self.progress)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet(f"color: {C_TEXT_SECONDARY}; font-size: 12px;")
        self.status_label.setVisible(False)
        layout.addWidget(self.status_label)

        layout.addStretch()

        scroll.setWidget(container)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)

    def _style_chip(self, cb):
        cb.setStyleSheet(f"""
            QCheckBox {{
                background: white;
                color: {C_TEXT_SECONDARY};
                border: 2px solid {C_BORDER};
                border-radius: 18px;
                padding: 7px 12px;
                font-size: 13px;
                font-weight: 700;
            }}
            QCheckBox:hover {{
                border-color: {C_PRIMARY};
                color: {C_PRIMARY};
                background: #F8FFF3;
            }}
            QCheckBox:checked {{
                background: {C_PRIMARY};
                color: white;
                border-color: {C_PRIMARY};
            }}
            QCheckBox::indicator {{
                width: 0px;
                height: 0px;
            }}
        """)

    def _browse_output(self):
        d = QFileDialog.getExistingDirectory(self, "选择保存目录", self.output_dir_edit.text())
        if d:
            self.output_dir_edit.setText(d)

    def _get_extensions(self):
        exts = set()

        if self.attachment_ext_presets.isChecked():
            exts.add(".pdf")

        if self.attachment_ext_doc.isChecked():
            exts.update([".doc", ".docx"])

        if self.attachment_ext_xls.isChecked():
            exts.update([".xls", ".xlsx", ".csv"])

        if self.attachment_ext_zip.isChecked():
            exts.update([".zip", ".rar", ".7z"])

        if self.attachment_ext_img.isChecked():
            exts.update([".png", ".jpg", ".jpeg", ".gif", ".bmp"])

        custom = self.attachment_ext_custom.text().strip()
        if custom:
            for e in custom.split(","):
                e = e.strip().lower()
                if e and not e.startswith("."):
                    e = "." + e
                if e:
                    exts.add(e)

        return list(exts) if exts else []

    def _get_conditions(self):
        date_from, date_to = self.date_picker.get_range()

        return {
            "title_mode": self.title_mode.value(),
            "title_text": self.title_text.text(),
            "body_mode": self.body_mode.value(),
            "body_text": self.body_text.text(),
            "sender_text": self.sender_text.text(),
            "recipient_text": self.recipient_text.text(),
            "date_mode": self.date_picker.value(),
            "date_from": date_from,
            "date_to": date_to,
            "has_attachments": self.has_attachment.isChecked(),
            "attachment_exts": self._get_extensions(),
            "auto_unzip": self.auto_unzip.isChecked(),
        }

    def _get_action(self):
        if self.action_attachment.isChecked():
            return "attachment"
        if self.action_email.isChecked():
            return "email"
        return "both"

    def _on_source_changed(self, value):
        self.imap_panel.setVisible(value == "imap")
        subtitle = self.findChild(QLabel, "subtitle")
        if value == "outlook":
            self.findChild(QLabel, "subtitle").setText("从 Outlook 自动收集邮件、附件和压缩包")
        else:
            self.findChild(QLabel, "subtitle").setText("从 IMAP 邮箱自动收集邮件、附件和压缩包")

    def _get_imap_config(self):
        if self.mail_source.value() != "imap":
            return None
        return {
            "server": self.imap_server.text().strip(),
            "port": int(self.imap_port.text().strip() or "993"),
            "account": self.imap_account.text().strip(),
            "password": self.imap_password.text(),
        }

    def _run(self):
        output = self.output_dir_edit.text().strip()

        if not output:
            QMessageBox.warning(self, "提示", "请选择保存目录")
            return

        os.makedirs(output, exist_ok=True)

        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.status_label.setVisible(True)
        self.status_label.setText("正在初始化...")

        conditions = refresh_dynamic_dates(self._get_conditions())

        if self.mail_source.value() == "imap":
            imap_config = self._get_imap_config()
            if not imap_config or not imap_config["server"] or not imap_config["account"] or not imap_config["password"]:
                QMessageBox.warning(self, "提示", "请填写完整的 IMAP 配置和本次运行密码")
                return
            self.worker = IMAPWorker(conditions, output, self._get_action(), imap_config)
        else:
            self.worker = OutlookWorker(conditions, output, self._get_action())

        self.worker.progress.connect(self._on_progress)
        self.worker.finished.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    def _on_progress(self, msg):
        self.status_label.setText(msg)

    def _on_finished(self, result):
        self.progress.setRange(0, 100)
        self.progress.setValue(100)
        self.status_label.setText("✅ 完成!")

        conditions = result.get("conditions", {})
        fail_reasons = result.get("fail_reasons", {})

        msg = "━━━ 搜索条件 ━━━\n"
        msg += f"标题: {conditions.get('title_text') or '(未设置)'}  [{conditions.get('title_mode', 'contains')}]\n"
        msg += f"正文: {conditions.get('body_text') or '(未设置)'}  [{conditions.get('body_mode', 'any')}]\n"
        msg += f"发件人: {conditions.get('sender_text') or '(未设置)'}\n"
        msg += f"收件人: {conditions.get('recipient_text') or '(未设置)'}\n"

        if conditions.get("date_from"):
            msg += f"日期从: {conditions['date_from']}\n"
        else:
            msg += "日期从: 不限\n"

        if conditions.get("date_to"):
            msg += f"日期到: {conditions['date_to']}\n"
        else:
            msg += "日期到: 不限\n"

        msg += f"仅附件: {'是' if conditions.get('has_attachments') else '否'}\n"
        msg += f"附件类型: {', '.join(conditions.get('attachment_exts', [])) or '(不限)'}\n"
        msg += f"自动解压: {'是' if conditions.get('auto_unzip') else '否'}\n"

        msg += "\n━━━ 扫描结果 ━━━\n"
        msg += f"扫描邮件: {result['scanned']} 封\n"
        msg += f"匹配邮件: {result['matched']} 封\n"
        msg += f"保存文件: {result['saved']} 个\n"
        msg += f"跳过重复: {result.get('skipped_duplicate', 0)} 个\n"
        msg += f"保存位置: {result['output']}\n"

        if fail_reasons:
            msg += "\n━━━ 不匹配原因统计 ━━━\n"
            for r, cnt in sorted(fail_reasons.items(), key=lambda x: -x[1]):
                msg += f"  {r}: {cnt} 封\n"

        if result["errors"]:
            msg += f"\n⚠ 错误: {len(result['errors'])} 个\n" + "\n".join(result["errors"][:3])

        QMessageBox.information(self, "任务完成", msg)

        self.progress.setVisible(False)
        self.status_label.setVisible(False)

    def _on_error(self, err):
        self.progress.setVisible(False)
        self.status_label.setVisible(False)
        QMessageBox.critical(self, "错误", f"执行失败:\n{err}")

    def _save_plan(self):
        editing = bool(self.editing_plan_id)

        dlg = SavePlanDialog(
            default_name=self.editing_plan_name if editing else "",
            editing=editing,
            parent=self
        )

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        mode, name = dlg.result_data()

        if not name:
            QMessageBox.warning(self, "提示", "请输入方案名称")
            return

        plans = load_plans()

        old_schedule = {
            "enabled": False,
            "time": "09:00",
            "months": 1,
            "end_date": "",
            "archive_mode": "by_date",
            "last_run_date": ""
        }
        old_created = datetime.now().strftime("%Y-%m-%d %H:%M")

        if self.editing_plan_id:
            for p in plans:
                if p.get("id") == self.editing_plan_id:
                    old_schedule = p.get("schedule", old_schedule)
                    old_created = p.get("created", old_created)
                    break

        data = {
            "id": self.editing_plan_id if mode == "overwrite" and self.editing_plan_id else new_plan_id(),
            "name": name,
            "conditions": self._get_conditions(),
            "action": self._get_action(),
            "output_dir": self.output_dir_edit.text(),
            "mail_source": self.mail_source.value(),
            "imap_config": self._get_imap_config() if self.mail_source.value() == "imap" else None,
            "created": old_created if mode == "overwrite" and self.editing_plan_id else datetime.now().strftime("%Y-%m-%d %H:%M"),
            "updated": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "schedule": old_schedule if mode == "overwrite" and self.editing_plan_id else {
                "enabled": False,
                "time": "09:00",
                "months": 1,
                "end_date": "",
                "archive_mode": "by_date",
                "last_run_date": ""
            }
        }

        if mode == "overwrite" and self.editing_plan_id:
            plans = [data if p.get("id") == self.editing_plan_id else p for p in plans]
            QMessageBox.information(self, "已覆盖", f"方案「{name}」已覆盖保存。")
        else:
            plans.append(data)
            QMessageBox.information(self, "已保存", f"方案「{name}」已另存为新方案。")

        save_plans(plans)

        self.editing_plan_id = data["id"]
        self.editing_plan_name = name

        self.main.plans_page.refresh()

    def load_plan_for_edit(self, plan):
        self.editing_plan_id = plan.get("id")
        self.editing_plan_name = plan.get("name", "")

        c = plan.get("conditions", {})

        self.title_mode.set_value(
            "startswith" if c.get("title_mode") == "startswith" else "contains"
        )
        self.title_text.setText(c.get("title_text", ""))

        self.body_mode.set_value(
            "all" if c.get("body_mode") == "all" else "any"
        )
        self.body_text.setText(c.get("body_text", ""))

        self.sender_text.setText(c.get("sender_text", ""))
        self.recipient_text.setText(c.get("recipient_text", ""))

        date_mode = c.get("date_mode")
        if date_mode and date_mode != "custom":
            self.date_picker.set_value(date_mode)
        else:
            self.date_picker.set_range(c.get("date_from"), c.get("date_to"))

        self.has_attachment.setChecked(c.get("has_attachments", True))
        self.output_dir_edit.setText(plan.get("output_dir", str(DEFAULT_OUTPUT)))

        exts = c.get("attachment_exts", [])
        exts_lower = [e.lower() for e in exts]

        self.attachment_ext_presets.setChecked(".pdf" in exts_lower)
        self.attachment_ext_doc.setChecked(any(e in exts_lower for e in [".doc", ".docx"]))
        self.attachment_ext_xls.setChecked(any(e in exts_lower for e in [".xls", ".xlsx", ".csv"]))
        self.attachment_ext_zip.setChecked(any(e in exts_lower for e in [".zip", ".rar", ".7z"]))
        self.attachment_ext_img.setChecked(any(e in exts_lower for e in [".png", ".jpg", ".jpeg", ".gif", ".bmp"]))

        preset_all = {
            ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".csv",
            ".zip", ".rar", ".7z", ".png", ".jpg", ".jpeg",
            ".gif", ".bmp"
        }

        custom = [e for e in exts if e.lower() not in preset_all]
        self.attachment_ext_custom.setText(", ".join(custom))

        self.auto_unzip.setChecked(c.get("auto_unzip", True))

        act = plan.get("action", "both")
        if act == "attachment":
            self.action_attachment.setChecked(True)
        elif act == "email":
            self.action_email.setChecked(True)
        else:
            self.action_both.setChecked(True)

        # 还原邮箱来源和IMAP配置
        source = plan.get("mail_source", "outlook")
        self.mail_source.set_value(source)
        self.imap_panel.setVisible(source == "imap")
        if source == "imap":
            imap_cfg = plan.get("imap_config") or {}
            self.imap_server.setText(imap_cfg.get("server", ""))
            self.imap_port.setText(str(imap_cfg.get("port", "993")))
            self.imap_account.setText(imap_cfg.get("account", ""))
            self.imap_password.clear()

    def clear_editing_plan(self):
        self.editing_plan_id = None
        self.editing_plan_name = ""


# ─── 页面2: 我的方案 ───

class PlansPage(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main = main_window

        layout = QVBoxLayout(self)
        layout.setSpacing(18)
        layout.setContentsMargins(40, 30, 40, 30)

        header = QHBoxLayout()

        icon = QLabel("📚")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setFixedSize(58, 58)
        icon.setStyleSheet("""
            QLabel {
                background: #E8F5FF;
                border-radius: 29px;
                font-size: 30px;
                border: none;
            }
        """)
        header.addWidget(icon)

        text_box = QVBoxLayout()
        text_box.setSpacing(4)
        text_box.addWidget(TitleLabel("我的方案"))
        text_box.addWidget(SubtitleLabel("保存常用筛选规则，下次一键执行，也可以设置定时运行"))
        header.addLayout(text_box)
        header.addStretch()

        layout.addLayout(header)

        self.plan_list = QListWidget()
        self.plan_list.setSpacing(10)
        self.plan_list.setUniformItemSizes(False)
        self.plan_list.itemClicked.connect(self._on_click)
        layout.addWidget(self.plan_list)

        btn_row = QHBoxLayout()
        btn_row.addStretch()

        run_btn = GreenButton("🚀 执行选中方案")
        run_btn.clicked.connect(self._run_selected)
        btn_row.addWidget(run_btn)

        del_btn = DangerButton("🗑 删除选中")
        del_btn.clicked.connect(self._delete_selected)
        btn_row.addWidget(del_btn)

        layout.addLayout(btn_row)

        self.refresh()

    def refresh(self):
        self.plan_list.clear()
        self.plan_list.setSpacing(10)
        self.plan_list.setUniformItemSizes(False)

        plans = load_plans()

        if not plans:
            item = QListWidgetItem()
            item.setFlags(Qt.ItemFlag.NoItemFlags)
            item.setSizeHint(QSize(0, 96))

            widget = QWidget()
            widget.setStyleSheet(f"""
                QWidget {{
                    background: {C_CARD};
                    border-radius: 18px;
                }}
            """)

            wl = QVBoxLayout(widget)
            wl.setContentsMargins(18, 14, 18, 14)
            wl.setSpacing(6)

            empty_title = QLabel("还没有保存方案")
            empty_title.setMinimumHeight(26)
            empty_title.setFont(QFont("Microsoft YaHei UI", 14, QFont.Bold))
            empty_title.setStyleSheet(f"""
                QLabel {{
                    color: {C_TEXT};
                    background: transparent;
                    border: none;
                }}
            """)
            wl.addWidget(empty_title)

            empty_subtitle = QLabel("在「邮件收集助手」页面设置筛选条件后，可以保存为方案。")
            empty_subtitle.setMinimumHeight(22)
            empty_subtitle.setFont(QFont("Microsoft YaHei UI", 11))
            empty_subtitle.setStyleSheet(f"""
                QLabel {{
                    color: {C_TEXT_SECONDARY};
                    background: transparent;
                    border: none;
                }}
            """)
            wl.addWidget(empty_subtitle)

            self.plan_list.addItem(item)
            self.plan_list.setItemWidget(item, widget)
            return

        for plan in plans:
            item = QListWidgetItem()
            item.setData(Qt.ItemDataRole.UserRole, plan)
            item.setSizeHint(QSize(0, 118))

            widget = QWidget()
            widget.setStyleSheet(f"""
                QWidget {{
                    background: {C_CARD};
                    border-radius: 18px;
                }}
            """)

            outer = QHBoxLayout(widget)
            outer.setContentsMargins(18, 13, 14, 13)
            outer.setSpacing(10)

            # 左侧：名称 + 摘要
            left_box = QVBoxLayout()
            left_box.setSpacing(6)

            name_edit = QLineEdit(plan.get("name", "未命名方案"))
            name_edit.setMinimumHeight(34)
            name_edit.setFont(QFont("Microsoft YaHei UI", 14, QFont.Bold))
            name_edit.setToolTip("点击即可修改方案名称，离开输入框后自动保存")
            name_edit.setStyleSheet(f"""
                QLineEdit {{
                    color: {C_TEXT};
                    background: transparent;
                    border: none;
                    padding: 0px;
                }}
                QLineEdit:focus {{
                    background: white;
                    border: 2px solid {C_PRIMARY};
                    border-radius: 10px;
                    padding: 3px 8px;
                }}
            """)
            name_edit.editingFinished.connect(
                lambda p_id=plan.get("id"), edit=name_edit: self._rename_plan(p_id, edit.text())
            )
            left_box.addWidget(name_edit)

            action_text = action_to_text(plan.get("action"))
            schedule = plan.get("schedule", {})

            if schedule.get("enabled"):
                schedule_text = f"每天 {schedule.get('time', '09:00')} 执行，至 {schedule.get('end_date', '')}"
            else:
                schedule_text = "未定时"

            detail = f"处理方式：{action_text}  ｜  创建：{plan.get('created', '')}  ｜  定时：{schedule_text}"

            detail_lbl = QLabel(detail)
            detail_lbl.setMinimumHeight(24)
            detail_lbl.setFont(QFont("Microsoft YaHei UI", 11))
            detail_lbl.setStyleSheet(f"""
                QLabel {{
                    color: {C_TEXT_SECONDARY};
                    background: transparent;
                    border: none;
                }}
            """)
            left_box.addWidget(detail_lbl)

            outer.addLayout(left_box, 1)

            # 右侧按钮：明细 / 编辑 / 定时
            right_box = QHBoxLayout()
            right_box.setSpacing(6)
            right_box.setAlignment(Qt.AlignmentFlag.AlignCenter)

            detail_btn = SmallActionButton("明细", C_ACCENT_BLUE)
            detail_btn.clicked.connect(lambda checked=False, p=plan: self._show_detail(p))
            right_box.addWidget(detail_btn)

            edit_btn = SmallActionButton("编辑", C_PRIMARY)
            edit_btn.clicked.connect(lambda checked=False, p=plan: self._edit_plan(p))
            right_box.addWidget(edit_btn)

            schedule_btn = QPushButton("⏰ 设置定时")
            schedule_btn.setFixedSize(92, 34)
            schedule_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            schedule_btn.setFont(QFont("Microsoft YaHei UI", 10, QFont.Bold))
            schedule_btn.setStyleSheet(f"""
                QPushButton {{
                    background: white;
                    color: {C_ACCENT_ORANGE};
                    border: 2px solid {C_ACCENT_ORANGE};
                    border-radius: 12px;
                    padding: 0px;
                }}
                QPushButton:hover {{
                    background: #FFF4D6;
                    color: #C87500;
                }}
                QPushButton:pressed {{
                    background: {C_ACCENT_ORANGE};
                    color: white;
                }}
            """)

            schedule_btn.clicked.connect(lambda checked=False, p=plan: self._schedule_plan(p))
            right_box.addWidget(schedule_btn)

            outer.addLayout(right_box)

            self.plan_list.addItem(item)
            self.plan_list.setItemWidget(item, widget)

    def _rename_plan(self, plan_id, new_name):
        new_name = new_name.strip()

        if not new_name:
            return

        plans = load_plans()
        changed = False

        for p in plans:
            if p.get("id") == plan_id and p.get("name") != new_name:
                p["name"] = new_name
                p["updated"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                changed = True
                break

        if changed:
            save_plans(plans)
            self.refresh()

    def _show_detail(self, plan):
        # 重新读取一次，防止名称刚刚被修改但 plan 对象还是旧的
        fresh_plan = plan
        for p in load_plans():
            if p.get("id") == plan.get("id"):
                fresh_plan = p
                break

        dlg = PlanDetailDialog(fresh_plan, self)
        dlg.exec()

    def _edit_plan(self, plan):
        fresh_plan = plan
        for p in load_plans():
            if p.get("id") == plan.get("id"):
                fresh_plan = p
                break

        self.main.create_page.load_plan_for_edit(fresh_plan)
        self.main.nav.setCurrentIndex(0)
        self.main.navbar.set_active(0)

        QMessageBox.information(
            self,
            "已进入编辑",
            "方案已加载到「收集助手」。\n\n修改后点击「保存为方案」，可以选择覆盖原方案或另存为新方案。"
        )

    def _schedule_plan(self, plan):
        fresh_plan = plan
        for p in load_plans():
            if p.get("id") == plan.get("id"):
                fresh_plan = p
                break

        dlg = ScheduleDialog(fresh_plan, self)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        schedule_data = dlg.result_data()

        plans = load_plans()
        for p in plans:
            if p.get("id") == fresh_plan.get("id"):
                p["schedule"] = schedule_data
                p["updated"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                break

        save_plans(plans)
        self.refresh()

        if schedule_data.get("enabled"):
            QMessageBox.information(
                self,
                "定时已启用",
                f"方案「{fresh_plan.get('name')}」将每天 {schedule_data.get('time')} 自动执行。\n\n注意：程序需要保持打开。"
            )
        else:
            QMessageBox.information(
                self,
                "定时已关闭",
                f"方案「{fresh_plan.get('name')}」已关闭定时执行。"
            )

    def _on_click(self, item):
        pass

    def _run_selected(self):
        items = self.plan_list.selectedItems()

        if not items:
            QMessageBox.warning(self, "提示", "请先选择一个方案")
            return

        plan = items[0].data(Qt.ItemDataRole.UserRole)

        if not plan:
            return

        self.main.run_plan(plan, scheduled=False)

    def _delete_selected(self):
        items = self.plan_list.selectedItems()

        if not items:
            QMessageBox.warning(self, "提示", "请先选择一个方案")
            return

        plan = items[0].data(Qt.ItemDataRole.UserRole)

        if not plan:
            return

        reply = QMessageBox.question(
            self,
            "确认删除",
            f"删除方案「{plan.get('name', '')}」？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            plans = load_plans()
            plans = [p for p in plans if p.get("id") != plan.get("id")]
            save_plans(plans)
            self.refresh()


# ─── 底部导航栏 ───

class NavBar(QWidget):
    def __init__(self, page_names, on_changed):
        super().__init__()

        self.setFixedHeight(74)
        self.setStyleSheet(f"""
            QWidget {{
                background: {C_BG};
                border-top: 1px solid {C_BORDER};
            }}
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.buttons = []

        for i, (icon, name) in enumerate(page_names):
            btn = QPushButton(f"{icon}\n{name}")
            btn.setFont(QFont("Microsoft YaHei UI", 11))
            btn.setFixedHeight(74)
            btn.clicked.connect(lambda checked=False, idx=i: on_changed(idx))
            layout.addWidget(btn)
            self.buttons.append(btn)

        self.set_active(0)

    def set_active(self, index):
        for i, btn in enumerate(self.buttons):
            active = i == index
            color = C_PRIMARY if active else C_TEXT_SECONDARY
            border = C_PRIMARY if active else "transparent"
            weight = "800" if active else "600"

            btn.setStyleSheet(f"""
                QPushButton {{
                    background: transparent;
                    color: {color};
                    border: none;
                    border-top: 4px solid {border};
                    font-weight: {weight};
                    font-size: 11px;
                    border-radius: 0px;
                    padding-top: 6px;
                }}
                QPushButton:hover {{
                    color: {C_PRIMARY};
                    background: #FAFAFA;
                }}
            """)


# ─── 主窗口 ───

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("MailCollector - 邮件附件收集器")
        self.setMinimumSize(860, 680)
        self.resize(1040, 780)

        self.setStyleSheet(STYLE)

        central = QWidget()
        self.setCentralWidget(central)

        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        self.nav = QStackedWidget()

        self.create_page = CreateTaskPage(self)
        self.plans_page = PlansPage(self)

        self.nav.addWidget(self.create_page)
        self.nav.addWidget(self.plans_page)

        main_layout.addWidget(self.nav)

        self.navbar = NavBar(
            [("🦉", "收集助手"), ("📚", "我的方案")],
            self._on_nav
        )
        main_layout.addWidget(self.navbar)

        # 定时任务相关
        self.scheduled_workers = []

        self.schedule_timer = QTimer(self)
        self.schedule_timer.timeout.connect(self._check_schedules)
        self.schedule_timer.start(30000)  # 每 30 秒检查一次

    def _on_nav(self, index):
        self.nav.setCurrentIndex(index)
        self.navbar.set_active(index)

        if index == 1:
            self.plans_page.refresh()

    def run_plan(self, plan, scheduled=False):
        """
        统一执行方案。
        scheduled=False：手动执行，会跳回收集助手页面并使用页面进度条。
        scheduled=True：定时执行，后台执行并去重。
        """
        fresh_plan = plan
        for p in load_plans():
            if p.get("id") == plan.get("id"):
                fresh_plan = p
                break

        output = fresh_plan.get("output_dir", str(DEFAULT_OUTPUT))
        os.makedirs(output, exist_ok=True)

        conditions = refresh_dynamic_dates(fresh_plan.get("conditions", {}))
        action = fresh_plan.get("action", "both")

        if scheduled:
            schedule = fresh_plan.get("schedule", {})

            archive_current_files(
                output,
                schedule.get("archive_mode", "by_date")
            )

            conditions["_dedupe"] = True
            conditions["_dedupe_root"] = output

            source = fresh_plan.get("mail_source", "outlook")
            if source == "imap":
                imap_config = resolve_imap_config(fresh_plan.get("imap_config"))
                if imap_config.get("server") and imap_config.get("account") and imap_config.get("password"):
                    worker = IMAPWorker(conditions, output, action, imap_config)
                else:
                    print("[定时跳过] IMAP 方案缺少运行密码；请设置 MAILCOLLECTOR_IMAP_PASSWORD")
                    return
            else:
                worker = OutlookWorker(conditions, output, action)
            self.scheduled_workers.append(worker)

            worker.finished.connect(
                lambda result, w=worker, p=fresh_plan: self._on_scheduled_finished(result, w, p)
            )
            worker.error.connect(
                lambda err, w=worker, p=fresh_plan: self._on_scheduled_error(err, w, p)
            )

            worker.start()

        else:
            self.create_page.load_plan_for_edit(fresh_plan)
            self.nav.setCurrentIndex(0)
            self.navbar.set_active(0)
            self.create_page._run()

    def _check_schedules(self):
        now = datetime.now()
        today_str = now.strftime("%Y-%m-%d")
        current_time = now.strftime("%H:%M")

        plans = load_plans()
        changed = False

        for plan in plans:
            schedule = plan.get("schedule", {})

            if not schedule.get("enabled"):
                continue

            end_date = schedule.get("end_date", "")
            if end_date and today_str > end_date:
                schedule["enabled"] = False
                changed = True
                continue

            # 今天已经执行过
            if schedule.get("last_run_date") == today_str:
                continue

            # 时间没到
            if schedule.get("time") != current_time:
                continue

            # 标记今天已经执行，避免同一分钟内重复执行
            schedule["last_run_date"] = today_str
            changed = True

            self.run_plan(plan, scheduled=True)

        if changed:
            save_plans(plans)
            self.plans_page.refresh()

    def _on_scheduled_finished(self, result, worker, plan):
        try:
            self.scheduled_workers.remove(worker)
        except ValueError:
            pass

        # 定时任务完成后静默处理，不弹窗打扰
        print(
            f"[定时完成] {plan.get('name', '')} | "
            f"匹配 {result.get('matched', 0)} 封 | "
            f"保存 {result.get('saved', 0)} 个 | "
            f"跳过重复 {result.get('skipped_duplicate', 0)} 个"
        )

    def _on_scheduled_error(self, err, worker, plan):
        try:
            self.scheduled_workers.remove(worker)
        except ValueError:
            pass

        QMessageBox.warning(
            self,
            "定时任务失败",
            f"方案「{plan.get('name', '')}」执行失败：\n{err}"
        )


# ─── 入口 ───

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()

