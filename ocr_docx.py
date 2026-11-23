# -*- coding: utf-8 -*-
"""
PDF → Word (.docx): текстовый слой + OCR для сканов.

Движки распознавания:
  * Windows OCR (Windows.Media.Ocr) — офлайн, умеет русский. Основной.
  * Tesseract — если установлен; на этой машине только eng.

Рендер страниц PDF в картинки — через QtPdf (PySide6), отдельный бинарь не нужен.
"""

from __future__ import annotations

import asyncio
import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

from pypdf import PdfReader

# --------------------------------------------------------------------------- #
# Движки OCR
# --------------------------------------------------------------------------- #

ENGINE_WINDOWS = "windows"
ENGINE_TESSERACT = "tesseract"

TESSERACT_CANDIDATES = [
    r"C:\tools\Tesseract-OCR\tesseract.exe",
    r"C:\tools\tesseract\tesseract.exe",
    r"C:\tools\tesseract.exe",
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]


def find_tesseract() -> str | None:
    for path in TESSERACT_CANDIDATES:
        if os.path.isfile(path):
            return path
    found = shutil.which("tesseract")
    return found


def tesseract_languages(exe: str | None = None) -> list[str]:
    exe = exe or find_tesseract()
    if not exe:
        return []
    try:
        out = subprocess.run(
            [exe, "--list-langs"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:  # noqa: BLE001
        return []
    langs = []
    for line in (out.stdout or "").splitlines()[1:]:
        line = line.strip()
        if line and "\\" not in line and "/" not in line:
            langs.append(line)
    return langs


def windows_ocr_languages() -> list[str]:
    try:
        from winsdk.windows.media.ocr import OcrEngine

        return [lang.language_tag for lang in OcrEngine.available_recognizer_languages]
    except Exception:  # noqa: BLE001
        return []


def available_engines() -> list[tuple[str, str, list[str]]]:
    """[(ключ, человеческое имя, языки)] — только реально доступные движки."""
    result = []
    win_langs = windows_ocr_languages()
    if win_langs:
        result.append((ENGINE_WINDOWS, "Windows OCR (встроенный)", win_langs))
    exe = find_tesseract()
    if exe:
        result.append((ENGINE_TESSERACT, "Tesseract", tesseract_languages(exe)))
    return result


# --- Windows.Media.Ocr ------------------------------------------------------ #

async def _win_ocr_async(image_path: str, lang_tag: str) -> list[str]:
    from winsdk.windows.globalization import Language
    from winsdk.windows.graphics.imaging import BitmapDecoder
    from winsdk.windows.media.ocr import OcrEngine
    from winsdk.windows.storage import FileAccessMode, StorageFile

    engine = OcrEngine.try_create_from_language(Language(lang_tag))
    if engine is None:
        raise RuntimeError(
            f"Windows OCR не поддерживает язык «{lang_tag}». "
            "Добавьте языковой пакет в параметрах Windows."
        )
    storage_file = await StorageFile.get_file_from_path_async(image_path)
    stream = await storage_file.open_async(FileAccessMode.READ)
    decoder = await BitmapDecoder.create_async(stream)
    bitmap = await decoder.get_software_bitmap_async()
    result = await engine.recognize_async(bitmap)
    return [line.text for line in result.lines]


def windows_ocr(image_path: str, lang_tag: str) -> list[str]:
    loop = asyncio.new_event_loop()
    try:
        return loop.run_until_complete(_win_ocr_async(image_path, lang_tag))
    finally:
        loop.close()


# --- Tesseract -------------------------------------------------------------- #

def tesseract_ocr(image_path: str, lang: str, exe: str | None = None) -> list[str]:
    exe = exe or find_tesseract()
    if not exe:
        raise RuntimeError("Tesseract не найден.")
    proc = subprocess.run(
        [exe, image_path, "stdout", "-l", lang, "--psm", "3"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if proc.returncode != 0:
        raise RuntimeError(f"Tesseract: {(proc.stderr or '').strip()[:300]}")
    return (proc.stdout or "").splitlines()


def ocr_image(image_path: str, engine: str, lang: str) -> list[str]:
    if engine == ENGINE_TESSERACT:
        return tesseract_ocr(image_path, lang)
    return windows_ocr(image_path, lang)


# --------------------------------------------------------------------------- #
# Рендер страниц PDF
# --------------------------------------------------------------------------- #

class PageRenderer:
    """Ленивый рендер страниц PDF в PNG через QtPdf."""

    def __init__(self, pdf_path: Path, dpi: int = 200) -> None:
        from PySide6.QtPdf import QPdfDocument

        self._doc = QPdfDocument()
        status = self._doc.load(str(pdf_path))
        if self._doc.pageCount() == 0:
            raise RuntimeError(f"Не удалось открыть PDF для рендера ({status}).")
        self._dpi = dpi

    @property
    def page_count(self) -> int:
        return self._doc.pageCount()

    def render_to(self, index: int, out_path: Path) -> Path:
        from PySide6.QtCore import QSize

        size_pt = self._doc.pagePointSize(index)
        scale = self._dpi / 72.0
        width = max(1, int(size_pt.width() * scale))
        height = max(1, int(size_pt.height() * scale))
        image = self._doc.render(index, QSize(width, height))
        if image.isNull():
            raise RuntimeError(f"Пустой рендер страницы {index + 1}.")
        if not image.save(str(out_path), "PNG"):
            raise RuntimeError(f"Не удалось сохранить картинку страницы {index + 1}.")
        return out_path


# --------------------------------------------------------------------------- #
# Текст → абзацы
# --------------------------------------------------------------------------- #

SENTENCE_END = re.compile(r"[.!?:;»”\"')\]]\s*$")
TIMECODE_START = re.compile(r"^\s*\d{1,2}:\d{2}")


def lines_to_paragraphs(lines: list[str]) -> list[str]:
    """Склеивает строки распознавания в осмысленные абзацы."""
    paragraphs: list[str] = []
    buffer: list[str] = []

    def flush() -> None:
        if buffer:
            paragraphs.append(" ".join(buffer).strip())
            buffer.clear()

    for raw in lines:
        line = raw.strip()
        if not line:
            flush()
            continue
        if buffer and TIMECODE_START.match(line):
            flush()
        if buffer and buffer[-1].endswith("-"):
            buffer[-1] = buffer[-1][:-1] + line
        else:
            buffer.append(line)
        if SENTENCE_END.search(line):
            flush()
    flush()
    return [p for p in paragraphs if p]


def text_to_paragraphs(text: str) -> list[str]:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    blocks = re.split(r"\n\s*\n", text)
    result: list[str] = []
    for block in blocks:
        lines = [ln.strip() for ln in block.split("\n")]
        merged = lines_to_paragraphs(lines)
        result.extend(merged)
    return result


# --------------------------------------------------------------------------- #
# Конвертация
# --------------------------------------------------------------------------- #

MODE_AUTO = "auto"
MODE_OCR = "ocr"
MODE_TEXT = "text"

MIN_TEXT_CHARS = 25  # меньше этого на странице — считаем её сканом


@dataclass
class DocxOptions:
    mode: str = MODE_AUTO
    engine: str = ENGINE_WINDOWS
    lang: str = "ru"
    dpi: int = 300
    page_breaks: bool = True
    page_headings: bool = False
    font_name: str = "Calibri"
    font_size: float = 11.0


@dataclass
class DocxResult:
    source: Path
    output: Path
    pages: int
    ocr_pages: list[int] = field(default_factory=list)

    @property
    def summary(self) -> str:
        if not self.ocr_pages:
            return f"{self.source.name} → {self.output.name}: {self.pages} стр., текстовый слой"
        if len(self.ocr_pages) == self.pages:
            return f"{self.source.name} → {self.output.name}: {self.pages} стр., все через OCR"
        shown = ", ".join(str(p) for p in self.ocr_pages[:10])
        if len(self.ocr_pages) > 10:
            shown += "…"
        return (
            f"{self.source.name} → {self.output.name}: {self.pages} стр., "
            f"OCR на {len(self.ocr_pages)} ({shown})"
        )


def _new_document(opts: DocxOptions):
    from docx import Document
    from docx.shared import Pt

    document = Document()
    style = document.styles["Normal"]
    style.font.name = opts.font_name
    style.font.size = Pt(opts.font_size)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    for attr in (
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii",
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi",
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}cs",
    ):
        rfonts.set(attr, opts.font_name)
    paragraph_format = style.paragraph_format
    paragraph_format.space_after = Pt(opts.font_size * 0.5)
    return document


def pdf_to_docx(
    src: Path,
    dst: Path,
    opts: DocxOptions,
    progress=lambda value, text: None,
    should_stop=lambda: False,
) -> DocxResult:
    src = Path(src)
    dst = Path(dst)
    if not src.is_file():
        raise FileNotFoundError(f"Файл не найден: {src}")

    reader = PdfReader(str(src))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError(f"PDF защищён паролем: {src.name}") from exc

    total = len(reader.pages)
    document = _new_document(opts)
    ocr_pages: list[int] = []
    renderer: PageRenderer | None = None
    tmpdir = Path(tempfile.mkdtemp(prefix="pdf2docx_"))

    try:
        if opts.page_headings:
            document.add_heading(src.stem, level=1)

        for index in range(total):
            if should_stop():
                raise RuntimeError("Прервано пользователем.")

            progress(
                int(index / max(total, 1) * 95),
                f"{src.name}: страница {index + 1} из {total}",
            )

            paragraphs: list[str] = []
            layer_text = ""
            if opts.mode != MODE_OCR:
                try:
                    layer_text = reader.pages[index].extract_text() or ""
                except Exception:  # noqa: BLE001
                    layer_text = ""

            need_ocr = opts.mode == MODE_OCR or (
                opts.mode == MODE_AUTO and len(layer_text.strip()) < MIN_TEXT_CHARS
            )

            if need_ocr:
                if renderer is None:
                    renderer = PageRenderer(src, opts.dpi)
                image_path = tmpdir / f"p{index + 1:04d}.png"
                renderer.render_to(index, image_path)
                try:
                    lines = ocr_image(str(image_path), opts.engine, opts.lang)
                finally:
                    try:
                        image_path.unlink()
                    except OSError:
                        pass
                paragraphs = lines_to_paragraphs(lines)
                if paragraphs:
                    ocr_pages.append(index + 1)
                elif layer_text.strip():
                    paragraphs = text_to_paragraphs(layer_text)
            else:
                paragraphs = text_to_paragraphs(layer_text)

            if opts.page_headings:
                document.add_heading(f"Страница {index + 1}", level=2)

            if not paragraphs:
                document.add_paragraph("")
            for text in paragraphs:
                document.add_paragraph(text)

            if opts.page_breaks and index < total - 1:
                document.add_page_break()

        progress(97, f"{src.name}: запись Word")
        dst.parent.mkdir(parents=True, exist_ok=True)
        document.save(str(dst))
        progress(100, "Готово")
        return DocxResult(source=src, output=dst, pages=total, ocr_pages=ocr_pages)
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


def convert_many(
    files: list[Path],
    out_dir: Path,
    opts: DocxOptions,
    progress=lambda value, text: None,
    should_stop=lambda: False,
) -> list[DocxResult]:
    results: list[DocxResult] = []
    total = max(len(files), 1)
    for position, src in enumerate(files):
        base = position / total

        def scaled(value: int, text: str, _base=base) -> None:
            progress(int((_base + value / 100.0 / total) * 100), text)

        dst = Path(out_dir) / f"{Path(src).stem}.docx"
        results.append(pdf_to_docx(Path(src), dst, opts, scaled, should_stop))
    progress(100, "Готово")
    return results
