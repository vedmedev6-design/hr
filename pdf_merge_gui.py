# -*- coding: utf-8 -*-
"""
Склейка PDF + PDF → Word с OCR.

Вкладка 1 «Склейка PDF»  — объединяет PDF и текстовые файлы в один PDF.
Вкладка 2 «PDF → Word»   — конвертирует PDF в .docx; для сканов включается OCR
                            (встроенный в Windows движок или Tesseract).
"""

from __future__ import annotations

import html
import os
import re
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QObject, Qt, QThread, Signal
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ocr_docx  # noqa: E402

APP_NAME = "Склейка PDF"

TEXT_SUFFIXES = {".txt", ".md", ".log", ".csv", ".srt", ".vtt"}
PDF_SUFFIXES = {".pdf"}

PRIMARY_BUTTON_CSS = (
    "QPushButton { background: #1d4ed8; color: white; border: none;"
    " border-radius: 6px; font-weight: 600; }"
    "QPushButton:hover { background: #1e40af; }"
    "QPushButton:disabled { background: #9ca3af; }"
)

# Необязательный локальный файл со списком путей (по одному на строку),
# которые подставляются в список склейки при запуске. В репозиторий не попадает.
DEFAULTS_FILE = Path(__file__).resolve().parent / "defaults.txt"


def load_default_files() -> list[Path]:
    if not DEFAULTS_FILE.is_file():
        return []
    try:
        lines = DEFAULTS_FILE.read_text(encoding="utf-8").splitlines()
    except OSError:
        return []
    result = []
    for line in lines:
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        path = Path(line)
        if path.is_file():
            result.append(path)
    return result


# --------------------------------------------------------------------------- #
# Шрифты с кириллицей (для вёрстки текста в PDF)
# --------------------------------------------------------------------------- #

FONT_CANDIDATES = [
    ("PdfMergeSans", "arial.ttf", "arialbd.ttf", "ariali.ttf"),
    ("PdfMergeSans", "segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf"),
    ("PdfMergeSans", "calibri.ttf", "calibrib.ttf", "calibrii.ttf"),
    ("PdfMergeSans", "tahoma.ttf", "tahomabd.ttf", None),
    ("PdfMergeSans", "DejaVuSans.ttf", "DejaVuSans-Bold.ttf", "DejaVuSans-Oblique.ttf"),
]

FONT_DIRS = [
    Path(os.environ.get("SystemRoot", r"C:\Windows")) / "Fonts",
    Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Windows" / "Fonts",
]

_FONTS_READY = False


def _find_font_file(name: str) -> Path | None:
    for directory in FONT_DIRS:
        if not directory or not directory.is_dir():
            continue
        candidate = directory / name
        if candidate.is_file():
            return candidate
    return None


def register_fonts() -> tuple[str, str]:
    """Регистрирует TTF-шрифт с кириллицей. Возвращает (regular, bold)."""
    global _FONTS_READY
    if _FONTS_READY:
        return "PdfMergeSans", "PdfMergeSans-Bold"

    for family, regular, bold, italic in FONT_CANDIDATES:
        regular_path = _find_font_file(regular)
        if not regular_path:
            continue
        bold_path = _find_font_file(bold) if bold else None
        italic_path = _find_font_file(italic) if italic else None

        pdfmetrics.registerFont(TTFont(family, str(regular_path)))
        bold_name = f"{family}-Bold"
        if bold_path:
            pdfmetrics.registerFont(TTFont(bold_name, str(bold_path)))
        else:
            bold_name = family
        italic_name = f"{family}-Italic"
        if italic_path:
            pdfmetrics.registerFont(TTFont(italic_name, str(italic_path)))
        else:
            italic_name = family

        pdfmetrics.registerFontFamily(
            family, normal=family, bold=bold_name, italic=italic_name, boldItalic=bold_name
        )
        _FONTS_READY = True
        return family, bold_name

    _FONTS_READY = True
    return "Helvetica", "Helvetica-Bold"


# --------------------------------------------------------------------------- #
# Чтение текста
# --------------------------------------------------------------------------- #

def read_text(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            pass
    try:
        from charset_normalizer import from_bytes

        best = from_bytes(raw).best()
        if best is not None:
            return str(best)
    except Exception:  # noqa: BLE001
        pass
    for encoding in ("cp1251", "utf-16", "cp866"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


# --------------------------------------------------------------------------- #
# Вёрстка текста в PDF
# --------------------------------------------------------------------------- #

@dataclass
class TextOptions:
    font_size: float = 10.5
    page_size: str = "A4"
    title_page: bool = True
    highlight_speakers: bool = True


SPEAKER_RE = re.compile(
    r"^\s*(?P<time>\d{1,2}:\d{2}(?::\d{2})?(?:\s*[-–—]\s*\d{1,2}:\d{2}(?::\d{2})?)?)?"
    r"\s*(?P<who>[^:\n]{1,60}?)\s*:\s*(?P<rest>.*)$",
    re.DOTALL,
)


def _page_size(name: str):
    return letter if name.upper() == "LETTER" else A4


def _format_block(block: str, bold_font: str, highlight: bool) -> str:
    block = block.strip()
    if not highlight:
        return html.escape(block).replace("\n", "<br/>")

    match = SPEAKER_RE.match(block)
    if match and (match.group("time") or match.group("who")):
        time_part = (match.group("time") or "").strip()
        who_part = (match.group("who") or "").strip()
        rest = (match.group("rest") or "").strip()
        if who_part and len(who_part.split()) <= 6:
            head_bits = []
            if time_part:
                head_bits.append(f'<font color="#6b7280">{html.escape(time_part)}</font>')
            head_bits.append(f'<font name="{bold_font}">{html.escape(who_part)}:</font>')
            head = " ".join(head_bits)
            body = html.escape(rest).replace("\n", "<br/>")
            return f"{head} {body}" if body else head

    return html.escape(block).replace("\n", "<br/>")


def text_to_pdf(src: Path, dst: Path, opts: TextOptions) -> int:
    regular, bold = register_fonts()
    page = _page_size(opts.page_size)

    doc = SimpleDocTemplate(
        str(dst),
        pagesize=page,
        leftMargin=20 * mm,
        rightMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
        title=src.stem,
        author=APP_NAME,
    )

    body_style = ParagraphStyle(
        "body",
        fontName=regular,
        fontSize=opts.font_size,
        leading=opts.font_size * 1.45,
        spaceAfter=opts.font_size * 0.55,
        alignment=TA_JUSTIFY,
        textColor=colors.HexColor("#111827"),
        wordWrap="CJK",
    )
    title_style = ParagraphStyle(
        "title",
        fontName=bold,
        fontSize=opts.font_size + 6,
        leading=(opts.font_size + 6) * 1.3,
        spaceAfter=opts.font_size * 1.2,
        textColor=colors.HexColor("#0f172a"),
    )
    subtitle_style = ParagraphStyle(
        "subtitle",
        fontName=regular,
        fontSize=opts.font_size - 1,
        leading=(opts.font_size - 1) * 1.4,
        spaceAfter=opts.font_size * 1.6,
        textColor=colors.HexColor("#6b7280"),
    )

    raw = read_text(src).replace("\r\n", "\n").replace("\r", "\n")
    blocks = [b for b in re.split(r"\n\s*\n", raw) if b.strip()]

    story: list = []
    if opts.title_page:
        first = blocks[0].strip() if blocks else ""
        first_line = first.split("\n", 1)[0].strip()
        looks_like_title = (
            bool(first_line) and len(first_line) <= 120 and not SPEAKER_RE.match(first_line)
        )
        heading = first_line if looks_like_title else src.stem
        story.append(Paragraph(html.escape(heading), title_style))
        story.append(Paragraph(html.escape(src.name), subtitle_style))
        if looks_like_title:
            rest_of_first = first.split("\n", 1)[1] if "\n" in first else ""
            blocks = ([rest_of_first] if rest_of_first.strip() else []) + blocks[1:]

    for block in blocks:
        story.append(Paragraph(_format_block(block, bold, opts.highlight_speakers), body_style))

    if not story:
        story.append(Paragraph("(файл пуст)", body_style))
        story.append(Spacer(1, 1))

    doc.build(story)
    return len(PdfReader(str(dst)).pages)


# --------------------------------------------------------------------------- #
# Склейка
# --------------------------------------------------------------------------- #

@dataclass
class MergeResult:
    output: Path
    pages: int
    details: list[str]


def merge_files(
    files: list[Path],
    output: Path,
    opts: TextOptions,
    bookmarks: bool = True,
    progress=lambda value, text: None,
) -> MergeResult:
    if not files:
        raise ValueError("Не выбрано ни одного файла.")

    writer = PdfWriter()
    details: list[str] = []
    tmpdir = Path(tempfile.mkdtemp(prefix="pdfmerge_"))
    total = len(files)

    try:
        for index, src in enumerate(files, start=1):
            progress(int((index - 1) / total * 95), f"Обработка: {src.name}")
            if not src.is_file():
                raise FileNotFoundError(f"Файл не найден: {src}")

            suffix = src.suffix.lower()
            if suffix in PDF_SUFFIXES:
                reader = PdfReader(str(src))
                if reader.is_encrypted:
                    try:
                        reader.decrypt("")
                    except Exception as exc:  # noqa: BLE001
                        raise RuntimeError(f"PDF защищён паролем: {src.name}") from exc
                start = len(writer.pages)
                writer.append(reader)
                added = len(writer.pages) - start
                details.append(f"{src.name} — {added} стр. (PDF)")
            else:
                converted = tmpdir / f"{index:02d}_{src.stem}.pdf"
                added = text_to_pdf(src, converted, opts)
                start = len(writer.pages)
                writer.append(str(converted))
                details.append(f"{src.name} — {added} стр. (текст → PDF)")

            if bookmarks:
                writer.add_outline_item(title=src.stem[:120], page_number=start)

        progress(96, "Запись результата")
        output.parent.mkdir(parents=True, exist_ok=True)
        writer.add_metadata({"/Title": output.stem, "/Producer": APP_NAME})
        with open(output, "wb") as handle:
            writer.write(handle)
        pages = len(writer.pages)
        writer.close()
        progress(100, "Готово")
        return MergeResult(output=output, pages=pages, details=details)
    finally:
        for leftover in tmpdir.glob("*"):
            try:
                leftover.unlink()
            except OSError:
                pass
        try:
            tmpdir.rmdir()
        except OSError:
            pass


# --------------------------------------------------------------------------- #
# Потоки
# --------------------------------------------------------------------------- #

class Worker(QObject):
    progress = Signal(int, str)
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, job) -> None:
        super().__init__()
        self._job = job
        self._stop = False

    def stop(self) -> None:
        self._stop = True

    def should_stop(self) -> bool:
        return self._stop

    def run(self) -> None:
        try:
            result = self._job(
                lambda value, text: self.progress.emit(value, text),
                self.should_stop,
            )
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(str(exc))
        else:
            self.finished.emit(result)


class TaskRunner:
    """Запуск работы в фоновом QThread с единым управлением состоянием."""

    def __init__(self, owner: QWidget) -> None:
        self._owner = owner
        self.thread: QThread | None = None
        self.worker: Worker | None = None

    @property
    def busy(self) -> bool:
        return self.thread is not None

    def start(self, job, on_progress, on_finished, on_failed) -> None:
        self.thread = QThread(self._owner)
        self.worker = Worker(job)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(on_progress)
        self.worker.finished.connect(on_finished)
        self.worker.failed.connect(on_failed)
        self.thread.start()

    def stop_thread(self) -> None:
        if self.thread:
            self.thread.quit()
            self.thread.wait(5000)
        self.thread = None
        self.worker = None

    def request_stop(self) -> None:
        if self.worker:
            self.worker.stop()


# --------------------------------------------------------------------------- #
# Вкладка 1: склейка
# --------------------------------------------------------------------------- #

class MergeTab(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setAcceptDrops(True)
        self.runner = TaskRunner(self)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        hint = QLabel(
            "Добавьте файлы (.pdf, .txt) и расставьте их в нужном порядке — "
            "перетаскиванием или кнопками. Текст будет свёрстан в PDF с кириллицей."
        )
        hint.setWordWrap(True)
        hint.setEnabled(False)
        root.addWidget(hint)

        files_box = QGroupBox("Файлы (порядок = порядок в итоговом PDF)")
        files_layout = QHBoxLayout(files_box)

        self.list = QListWidget()
        self.list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.list.setDragDropMode(QAbstractItemView.InternalMove)
        self.list.setDefaultDropAction(Qt.MoveAction)
        files_layout.addWidget(self.list, 1)

        buttons = QVBoxLayout()
        buttons.setSpacing(6)
        self.btn_add = QPushButton("Добавить…")
        self.btn_up = QPushButton("Вверх")
        self.btn_down = QPushButton("Вниз")
        self.btn_remove = QPushButton("Удалить")
        self.btn_clear = QPushButton("Очистить")
        for btn in (self.btn_add, self.btn_up, self.btn_down, self.btn_remove, self.btn_clear):
            btn.setMinimumWidth(120)
            buttons.addWidget(btn)
        buttons.addStretch(1)
        files_layout.addLayout(buttons)
        root.addWidget(files_box, 1)

        opts_box = QGroupBox("Вёрстка текстовых файлов")
        opts_layout = QHBoxLayout(opts_box)
        opts_layout.addWidget(QLabel("Размер шрифта:"))
        self.font_size = QSpinBox()
        self.font_size.setRange(6, 20)
        self.font_size.setValue(11)
        self.font_size.setSuffix(" pt")
        opts_layout.addWidget(self.font_size)
        opts_layout.addSpacing(16)
        opts_layout.addWidget(QLabel("Формат страницы:"))
        self.page_size = QComboBox()
        self.page_size.addItems(["A4", "Letter"])
        opts_layout.addWidget(self.page_size)
        opts_layout.addSpacing(16)
        self.chk_title = QCheckBox("Заголовок файла")
        self.chk_title.setChecked(True)
        opts_layout.addWidget(self.chk_title)
        self.chk_speakers = QCheckBox("Выделять спикеров и таймкоды")
        self.chk_speakers.setChecked(True)
        opts_layout.addWidget(self.chk_speakers)
        opts_layout.addStretch(1)
        root.addWidget(opts_box)

        out_box = QGroupBox("Результат")
        out_layout = QVBoxLayout(out_box)
        out_row = QHBoxLayout()
        self.output = QLineEdit()
        self.output.setPlaceholderText("Путь к итоговому PDF")
        self.btn_output = QPushButton("Выбрать…")
        out_row.addWidget(self.output, 1)
        out_row.addWidget(self.btn_output)
        out_layout.addLayout(out_row)
        extra_row = QHBoxLayout()
        self.chk_bookmarks = QCheckBox("Закладки по исходным файлам")
        self.chk_bookmarks.setChecked(True)
        self.chk_open = QCheckBox("Открыть PDF после склейки")
        self.chk_open.setChecked(True)
        extra_row.addWidget(self.chk_bookmarks)
        extra_row.addWidget(self.chk_open)
        extra_row.addStretch(1)
        out_layout.addLayout(extra_row)
        root.addWidget(out_box)

        self.progress = QProgressBar()
        root.addWidget(self.progress)

        self.status = QLabel("Готов к работе.")
        self.status.setWordWrap(True)
        self.status.setTextInteractionFlags(Qt.TextSelectableByMouse)
        root.addWidget(self.status)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        root.addWidget(line)

        run_row = QHBoxLayout()
        run_row.addStretch(1)
        self.btn_merge = QPushButton("Объединить в PDF")
        self.btn_merge.setMinimumHeight(38)
        self.btn_merge.setMinimumWidth(190)
        self.btn_merge.setStyleSheet(PRIMARY_BUTTON_CSS)
        run_row.addWidget(self.btn_merge)
        root.addLayout(run_row)

        self.btn_add.clicked.connect(self.add_files)
        self.btn_up.clicked.connect(lambda: self.move_selected(-1))
        self.btn_down.clicked.connect(lambda: self.move_selected(1))
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_clear.clicked.connect(self.list.clear)
        self.btn_output.clicked.connect(self.choose_output)
        self.btn_merge.clicked.connect(self.start_merge)
        self.list.model().rowsInserted.connect(lambda *_: self.suggest_output())
        self.list.model().rowsRemoved.connect(lambda *_: self.suggest_output())

        self.prefill()

    # ---------------- список ----------------

    def prefill(self) -> None:
        added = load_default_files()
        if added:
            self.add_paths(added)
            self.status.setText(
                f"Подставлены файлы из defaults.txt ({len(added)}). Можно менять список."
            )

    def add_paths(self, paths: list[Path]) -> None:
        existing = {self.list.item(i).data(Qt.UserRole) for i in range(self.list.count())}
        for path in paths:
            key = str(path.resolve())
            if key in existing:
                continue
            suffix = path.suffix.lower()
            if suffix not in PDF_SUFFIXES and suffix not in TEXT_SUFFIXES:
                continue
            kind = "PDF" if suffix in PDF_SUFFIXES else "текст"
            item = QListWidgetItem(f"[{kind}]  {path.name}")
            item.setData(Qt.UserRole, key)
            item.setToolTip(key)
            self.list.addItem(item)
            existing.add(key)
        self.suggest_output()

    def current_paths(self) -> list[Path]:
        return [Path(self.list.item(i).data(Qt.UserRole)) for i in range(self.list.count())]

    def add_files(self) -> None:
        paths = self.current_paths()
        start_dir = str(paths[0].parent) if paths else ""
        selected, _ = QFileDialog.getOpenFileNames(
            self,
            "Выберите PDF и текстовые файлы",
            start_dir,
            "Документы (*.pdf *.txt *.md *.log *.csv *.srt *.vtt);;PDF (*.pdf);;Текст (*.txt);;Все файлы (*.*)",
        )
        if selected:
            self.add_paths([Path(p) for p in selected])

    def move_selected(self, delta: int) -> None:
        rows = sorted((self.list.row(i) for i in self.list.selectedItems()), reverse=delta > 0)
        for row in rows:
            target = row + delta
            if target < 0 or target >= self.list.count():
                continue
            item = self.list.takeItem(row)
            self.list.insertItem(target, item)
            item.setSelected(True)

    def remove_selected(self) -> None:
        for item in self.list.selectedItems():
            self.list.takeItem(self.list.row(item))
        self.suggest_output()

    def suggest_output(self) -> None:
        if self.output.text().strip():
            return
        paths = self.current_paths()
        if not paths:
            return
        base = paths[0]
        self.output.setText(str(base.parent / f"{base.stem} (объединено).pdf"))

    def choose_output(self) -> None:
        start = self.output.text().strip() or str(Path.home() / "объединено.pdf")
        selected, _ = QFileDialog.getSaveFileName(self, "Сохранить итоговый PDF", start, "PDF (*.pdf)")
        if selected:
            if not selected.lower().endswith(".pdf"):
                selected += ".pdf"
            self.output.setText(selected)

    # ---------------- drag & drop ----------------

    def dragEnterEvent(self, event) -> None:  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:  # noqa: N802
        paths = [Path(u.toLocalFile()) for u in event.mimeData().urls() if u.isLocalFile()]
        if paths:
            self.add_paths(paths)
            event.acceptProposedAction()

    # ---------------- запуск ----------------

    def start_merge(self) -> None:
        if self.runner.busy:
            return
        files = self.current_paths()
        if not files:
            QMessageBox.warning(self, APP_NAME, "Добавьте хотя бы один файл.")
            return

        out_text = self.output.text().strip()
        if not out_text:
            QMessageBox.warning(self, APP_NAME, "Укажите, куда сохранить итоговый PDF.")
            return
        output = Path(out_text)
        if output.suffix.lower() != ".pdf":
            output = output.with_suffix(".pdf")

        if str(output.resolve()) in {str(f.resolve()) for f in files}:
            QMessageBox.warning(
                self, APP_NAME, "Итоговый файл совпадает с одним из исходных. Выберите другое имя."
            )
            return

        if output.exists():
            answer = QMessageBox.question(
                self,
                APP_NAME,
                f"Файл уже существует:\n{output}\n\nПерезаписать?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return

        opts = TextOptions(
            font_size=float(self.font_size.value()),
            page_size=self.page_size.currentText(),
            title_page=self.chk_title.isChecked(),
            highlight_speakers=self.chk_speakers.isChecked(),
        )
        bookmarks = self.chk_bookmarks.isChecked()

        def job(progress, _should_stop):
            return merge_files(files, output, opts, bookmarks=bookmarks, progress=progress)

        self.set_busy(True)
        self.progress.setValue(0)
        self.status.setText("Запуск…")
        self.runner.start(job, self.on_progress, self.on_finished, self.on_failed)

    def set_busy(self, busy: bool) -> None:
        for widget in (
            self.btn_merge, self.btn_add, self.btn_up, self.btn_down,
            self.btn_remove, self.btn_clear, self.btn_output, self.list,
        ):
            widget.setEnabled(not busy)

    def on_progress(self, value: int, text: str) -> None:
        self.progress.setValue(value)
        self.status.setText(text)

    def on_finished(self, result: MergeResult) -> None:
        self.runner.stop_thread()
        self.set_busy(False)
        self.progress.setValue(100)
        details = "\n".join(f"  • {line}" for line in result.details)
        self.status.setText(f"Готово: {result.output.name} — {result.pages} стр.\n{details}")
        QMessageBox.information(
            self, APP_NAME, f"PDF собран: {result.pages} стр.\n\n{result.output}\n\n{details}"
        )
        if self.chk_open.isChecked():
            try:
                os.startfile(str(result.output))  # noqa: S606
            except Exception:  # noqa: BLE001
                pass

    def on_failed(self, message: str) -> None:
        self.runner.stop_thread()
        self.set_busy(False)
        self.progress.setValue(0)
        self.status.setText(f"Ошибка: {message}")
        QMessageBox.critical(self, APP_NAME, f"Не удалось собрать PDF:\n\n{message}")


# --------------------------------------------------------------------------- #
# Вкладка 2: PDF → Word с OCR
# --------------------------------------------------------------------------- #

LANG_LABELS = {
    "ru": "Русский",
    "en": "English",
    "en-US": "English",
    "uk": "Українська",
    "bs": "Bosanski",
    "rus": "Русский",
    "eng": "English",
    "ukr": "Українська",
}


class OcrTab(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setAcceptDrops(True)
        self.runner = TaskRunner(self)
        self.engines = ocr_docx.available_engines()

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        hint = QLabel(
            "PDF превращается в документ Word. Если в PDF есть текстовый слой — "
            "берётся он; страницы-сканы распознаются OCR прямо на этом компьютере."
        )
        hint.setWordWrap(True)
        hint.setEnabled(False)
        root.addWidget(hint)

        files_box = QGroupBox("PDF-файлы")
        files_layout = QHBoxLayout(files_box)
        self.list = QListWidget()
        self.list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        files_layout.addWidget(self.list, 1)
        buttons = QVBoxLayout()
        buttons.setSpacing(6)
        self.btn_add = QPushButton("Добавить…")
        self.btn_remove = QPushButton("Удалить")
        self.btn_clear = QPushButton("Очистить")
        for btn in (self.btn_add, self.btn_remove, self.btn_clear):
            btn.setMinimumWidth(120)
            buttons.addWidget(btn)
        buttons.addStretch(1)
        files_layout.addLayout(buttons)
        root.addWidget(files_box, 1)

        opts_box = QGroupBox("Распознавание")
        opts_layout = QVBoxLayout(opts_box)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Режим:"))
        self.mode = QComboBox()
        self.mode.addItem("Авто: текстовый слой, OCR только для сканов", ocr_docx.MODE_AUTO)
        self.mode.addItem("Всегда OCR (игнорировать текстовый слой)", ocr_docx.MODE_OCR)
        self.mode.addItem("Только текстовый слой (без OCR)", ocr_docx.MODE_TEXT)
        row1.addWidget(self.mode, 1)
        opts_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Движок OCR:"))
        self.engine = QComboBox()
        for key, title, _langs in self.engines:
            self.engine.addItem(title, key)
        if not self.engines:
            self.engine.addItem("OCR недоступен", "")
            self.engine.setEnabled(False)
        row2.addWidget(self.engine, 1)

        row2.addSpacing(12)
        row2.addWidget(QLabel("Язык:"))
        self.lang = QComboBox()
        row2.addWidget(self.lang, 1)

        row2.addSpacing(12)
        row2.addWidget(QLabel("Качество:"))
        self.dpi = QSpinBox()
        self.dpi.setRange(100, 450)
        self.dpi.setSingleStep(50)
        self.dpi.setValue(300)
        self.dpi.setSuffix(" dpi")
        row2.addWidget(self.dpi)
        opts_layout.addLayout(row2)

        row3 = QHBoxLayout()
        self.chk_breaks = QCheckBox("Разрыв страницы между страницами PDF")
        self.chk_breaks.setChecked(True)
        self.chk_headings = QCheckBox("Заголовки «Страница N»")
        self.chk_open = QCheckBox("Открыть после конвертации")
        self.chk_open.setChecked(True)
        row3.addWidget(self.chk_breaks)
        row3.addWidget(self.chk_headings)
        row3.addWidget(self.chk_open)
        row3.addStretch(1)
        opts_layout.addLayout(row3)
        root.addWidget(opts_box)

        out_box = QGroupBox("Куда сохранить")
        out_layout = QHBoxLayout(out_box)
        self.out_dir = QLineEdit()
        self.out_dir.setPlaceholderText("Папка для .docx (по умолчанию — рядом с PDF)")
        self.btn_out_dir = QPushButton("Выбрать…")
        out_layout.addWidget(self.out_dir, 1)
        out_layout.addWidget(self.btn_out_dir)
        root.addWidget(out_box)

        self.progress = QProgressBar()
        root.addWidget(self.progress)

        self.status = QLabel(self._engine_status())
        self.status.setWordWrap(True)
        self.status.setTextInteractionFlags(Qt.TextSelectableByMouse)
        root.addWidget(self.status)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        root.addWidget(line)

        run_row = QHBoxLayout()
        self.btn_cancel = QPushButton("Остановить")
        self.btn_cancel.setEnabled(False)
        run_row.addWidget(self.btn_cancel)
        run_row.addStretch(1)
        self.btn_convert = QPushButton("Конвертировать в Word")
        self.btn_convert.setMinimumHeight(38)
        self.btn_convert.setMinimumWidth(210)
        self.btn_convert.setStyleSheet(PRIMARY_BUTTON_CSS)
        run_row.addWidget(self.btn_convert)
        root.addLayout(run_row)

        self.btn_add.clicked.connect(self.add_files)
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_clear.clicked.connect(self.list.clear)
        self.btn_out_dir.clicked.connect(self.choose_out_dir)
        self.btn_convert.clicked.connect(self.start_convert)
        self.btn_cancel.clicked.connect(self.runner.request_stop)
        self.engine.currentIndexChanged.connect(self.reload_languages)
        self.mode.currentIndexChanged.connect(self.update_enabled)

        self.reload_languages()
        self.update_enabled()

    # ---------------- вспомогательное ----------------

    def _engine_status(self) -> str:
        if not self.engines:
            return (
                "OCR недоступен: не найден ни встроенный движок Windows, ни Tesseract. "
                "Конвертировать можно только PDF с текстовым слоем."
            )
        parts = []
        for _key, title, langs in self.engines:
            shown = ", ".join(LANG_LABELS.get(lang, lang) for lang in langs if lang != "osd")
            parts.append(f"{title}: {shown or 'нет языков'}")
        return "Доступно — " + "; ".join(parts)

    def reload_languages(self) -> None:
        key = self.engine.currentData()
        self.lang.clear()
        langs: list[str] = []
        for engine_key, _title, engine_langs in self.engines:
            if engine_key == key:
                langs = [lang for lang in engine_langs if lang != "osd"]
                break
        for lang in langs:
            self.lang.addItem(LANG_LABELS.get(lang, lang), lang)
        for index in range(self.lang.count()):
            if str(self.lang.itemData(index)).lower().startswith("ru"):
                self.lang.setCurrentIndex(index)
                break
        self.lang.setEnabled(self.lang.count() > 0)

    def update_enabled(self) -> None:
        ocr_used = self.mode.currentData() != ocr_docx.MODE_TEXT
        for widget in (self.engine, self.lang, self.dpi):
            widget.setEnabled(ocr_used and bool(self.engines))

    # ---------------- список ----------------

    def add_paths(self, paths: list[Path]) -> None:
        existing = {self.list.item(i).data(Qt.UserRole) for i in range(self.list.count())}
        for path in paths:
            if path.suffix.lower() not in PDF_SUFFIXES:
                continue
            key = str(path.resolve())
            if key in existing:
                continue
            item = QListWidgetItem(path.name)
            item.setData(Qt.UserRole, key)
            item.setToolTip(key)
            self.list.addItem(item)
            existing.add(key)

    def current_paths(self) -> list[Path]:
        return [Path(self.list.item(i).data(Qt.UserRole)) for i in range(self.list.count())]

    def add_files(self) -> None:
        paths = self.current_paths()
        start_dir = str(paths[0].parent) if paths else ""
        selected, _ = QFileDialog.getOpenFileNames(
            self, "Выберите PDF", start_dir, "PDF (*.pdf);;Все файлы (*.*)"
        )
        if selected:
            self.add_paths([Path(p) for p in selected])

    def remove_selected(self) -> None:
        for item in self.list.selectedItems():
            self.list.takeItem(self.list.row(item))

    def choose_out_dir(self) -> None:
        start = self.out_dir.text().strip()
        if not start:
            paths = self.current_paths()
            start = str(paths[0].parent) if paths else str(Path.home())
        selected = QFileDialog.getExistingDirectory(self, "Папка для .docx", start)
        if selected:
            self.out_dir.setText(selected)

    def dragEnterEvent(self, event) -> None:  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event) -> None:  # noqa: N802
        paths = [Path(u.toLocalFile()) for u in event.mimeData().urls() if u.isLocalFile()]
        if paths:
            self.add_paths(paths)
            event.acceptProposedAction()

    # ---------------- запуск ----------------

    def start_convert(self) -> None:
        if self.runner.busy:
            return
        files = self.current_paths()
        if not files:
            QMessageBox.warning(self, APP_NAME, "Добавьте хотя бы один PDF.")
            return

        mode = self.mode.currentData()
        if mode != ocr_docx.MODE_TEXT and not self.engines:
            QMessageBox.warning(
                self, APP_NAME, "OCR недоступен. Выберите режим «Только текстовый слой»."
            )
            return

        out_dir_text = self.out_dir.text().strip()
        out_dir = Path(out_dir_text) if out_dir_text else files[0].parent

        existing = [
            f"{Path(f).stem}.docx" for f in files if (out_dir / f"{Path(f).stem}.docx").exists()
        ]
        if existing:
            answer = QMessageBox.question(
                self,
                APP_NAME,
                "Уже существуют:\n"
                + "\n".join(f"  • {name}" for name in existing[:10])
                + "\n\nПерезаписать?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return

        opts = ocr_docx.DocxOptions(
            mode=mode,
            engine=self.engine.currentData() or ocr_docx.ENGINE_WINDOWS,
            lang=self.lang.currentData() or "ru",
            dpi=self.dpi.value(),
            page_breaks=self.chk_breaks.isChecked(),
            page_headings=self.chk_headings.isChecked(),
        )

        def job(progress, should_stop):
            return ocr_docx.convert_many(files, out_dir, opts, progress, should_stop)

        self.set_busy(True)
        self.progress.setValue(0)
        self.status.setText("Запуск…")
        self.runner.start(job, self.on_progress, self.on_finished, self.on_failed)

    def set_busy(self, busy: bool) -> None:
        for widget in (
            self.btn_convert, self.btn_add, self.btn_remove, self.btn_clear,
            self.btn_out_dir, self.list, self.mode,
        ):
            widget.setEnabled(not busy)
        self.btn_cancel.setEnabled(busy)
        if busy:
            for widget in (self.engine, self.lang, self.dpi):
                widget.setEnabled(False)
        else:
            self.update_enabled()

    def on_progress(self, value: int, text: str) -> None:
        self.progress.setValue(value)
        self.status.setText(text)

    def on_finished(self, results: list) -> None:
        self.runner.stop_thread()
        self.set_busy(False)
        self.progress.setValue(100)
        details = "\n".join(f"  • {r.summary}" for r in results)
        self.status.setText(f"Готово, файлов: {len(results)}\n{details}")
        QMessageBox.information(self, APP_NAME, f"Готово.\n\n{details}")
        if self.chk_open.isChecked() and results:
            try:
                os.startfile(str(results[0].output))  # noqa: S606
            except Exception:  # noqa: BLE001
                pass

    def on_failed(self, message: str) -> None:
        self.runner.stop_thread()
        self.set_busy(False)
        self.progress.setValue(0)
        self.status.setText(f"Ошибка: {message}")
        QMessageBox.critical(self, APP_NAME, f"Не удалось конвертировать:\n\n{message}")


# --------------------------------------------------------------------------- #
# Окно
# --------------------------------------------------------------------------- #

def _make_icon() -> QIcon:
    icon_file = Path(__file__).resolve().parent / "icon.ico"
    if icon_file.is_file():
        return QIcon(str(icon_file))
    pixmap = QPixmap(64, 64)
    pixmap.fill(QColor("transparent"))
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    painter.setBrush(QColor("#b91c1c"))
    painter.setPen(Qt.NoPen)
    painter.drawRoundedRect(6, 4, 52, 56, 8, 8)
    painter.setPen(QColor("white"))
    font = QFont()
    font.setBold(True)
    font.setPointSize(18)
    painter.setFont(font)
    painter.drawText(pixmap.rect(), Qt.AlignCenter, "PDF")
    painter.end()
    return QIcon(pixmap)


class MainWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setMinimumSize(840, 640)

        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        header = QLabel("PDF: склейка и конвертация в Word")
        header_font = header.font()
        header_font.setPointSize(header_font.pointSize() + 3)
        header_font.setBold(True)
        header.setFont(header_font)
        root.addWidget(header)

        self.tabs = QTabWidget()
        self.merge_tab = MergeTab()
        self.ocr_tab = OcrTab()
        self.tabs.addTab(self.merge_tab, "Склейка PDF")
        self.tabs.addTab(self.ocr_tab, "PDF → Word (OCR)")
        root.addWidget(self.tabs, 1)

    def closeEvent(self, event) -> None:  # noqa: N802
        for tab in (self.merge_tab, self.ocr_tab):
            tab.runner.request_stop()
            tab.runner.stop_thread()
        super().closeEvent(event)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def run_cli(argv: list[str]) -> int:
    if len(argv) < 2:
        print("Использование: --cli <итог.pdf> <файл1> [файл2 ...]")
        return 2
    output = Path(argv[0])
    files = [Path(a) for a in argv[1:]]
    result = merge_files(files, output, TextOptions(), progress=lambda v, t: print(f"[{v:3d}%] {t}"))
    print(f"\nГотово: {result.output} — {result.pages} стр.")
    for line in result.details:
        print(f"  • {line}")
    return 0


def run_docx_cli(argv: list[str]) -> int:
    if not argv:
        print(
            "Использование: --docx <файл.pdf> [ещё.pdf ...] "
            "[--out <папка>] [--mode auto|ocr|text] [--lang ru]"
        )
        return 2
    files: list[Path] = []
    out_dir: Path | None = None
    mode = ocr_docx.MODE_AUTO
    lang = "ru"
    index = 0
    while index < len(argv):
        arg = argv[index]
        if arg == "--out" and index + 1 < len(argv):
            out_dir = Path(argv[index + 1])
            index += 2
            continue
        if arg == "--mode" and index + 1 < len(argv):
            mode = argv[index + 1]
            index += 2
            continue
        if arg == "--lang" and index + 1 < len(argv):
            lang = argv[index + 1]
            index += 2
            continue
        files.append(Path(arg))
        index += 1

    if not files:
        print("Не указан ни один PDF.")
        return 2
    target = out_dir or files[0].parent
    opts = ocr_docx.DocxOptions(mode=mode, lang=lang)
    results = ocr_docx.convert_many(
        files, target, opts, progress=lambda v, t: print(f"[{v:3d}%] {t}")
    )
    print()
    for result in results:
        print(f"  • {result.summary}")
    return 0


def main() -> int:
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        return run_cli(sys.argv[2:])
    if len(sys.argv) > 1 and sys.argv[1] == "--docx":
        return run_docx_cli(sys.argv[2:])

    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setWindowIcon(_make_icon())
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
