from __future__ import annotations

import re
import tempfile
from pathlib import Path

import pymupdf

from library_core import extract_docx, extract_markdown


def _visual_width(char: str) -> float:
    if char == "\t":
        return 4.0
    return 1.0 if ord(char) > 255 else 0.56


def _wrap_line(line: str, maximum: float = 48.0) -> list[str]:
    if not line:
        return [""]
    output: list[str] = []
    current: list[str] = []
    width = 0.0
    for char in line.expandtabs(4):
        amount = _visual_width(char)
        if current and width + amount > maximum:
            output.append("".join(current).rstrip())
            current = []
            width = 0.0
        current.append(char)
        width += amount
    output.append("".join(current).rstrip())
    return output


def create_searchable_pdf(text: str, target: str | Path, title: str = "") -> Path:
    """Create a selectable/searchable A4 PDF with an embedded CJK text font."""
    target = Path(target)
    document = pymupdf.open()
    page = None
    y = 0.0
    line_height = 16.0
    top, bottom, left = 52.0, 52.0, 55.0

    def new_page():
        nonlocal page, y
        page = document.new_page(width=595, height=842)
        y = top

    new_page()
    source_lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if title:
        source_lines = [title, "", *source_lines]
    for source_line in source_lines:
        # Retain readable Markdown text while removing invisible control characters.
        clean = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", source_line)
        for line in _wrap_line(clean):
            if y + line_height > 842 - bottom:
                new_page()
            page.insert_text((left, y), line or " ", fontname="china-s", fontsize=10.5)
            y += line_height
    document.set_metadata({"title": title or target.stem, "creator": "智能文档学习库"})
    document.save(target, garbage=4, deflate=True)
    document.close()
    return target


def convert_document_to_searchable_pdf(source: str | Path, target: str | Path) -> Path:
    """Convert DOC/DOCX/Markdown to a searchable PDF, preferring Word for layout fidelity."""
    source = Path(source).resolve()
    target = Path(target).resolve()
    suffix = source.suffix.lower()
    target.parent.mkdir(parents=True, exist_ok=True)
    if suffix in {".doc", ".docx"}:
        try:
            import pythoncom
            import win32com.client

            pythoncom.CoInitialize()
            word = None
            document = None
            try:
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                document = word.Documents.Open(str(source), ReadOnly=True, AddToRecentFiles=False)
                document.ExportAsFixedFormat(str(target), 17)
                return target
            finally:
                if document is not None:
                    document.Close(False)
                if word is not None:
                    word.Quit()
                pythoncom.CoUninitialize()
        except Exception:
            if suffix == ".doc":
                raise ValueError("旧版 .doc 转换需要本机安装 Microsoft Word；可先另存为 .docx")
            chunks = extract_docx(source)
            return create_searchable_pdf("\n\n".join(chunk.content for chunk in chunks), target, source.stem)
    if suffix in {".md", ".markdown"}:
        chunks = extract_markdown(source)
        return create_searchable_pdf("\n\n".join(chunk.content for chunk in chunks), target, source.stem)
    raise ValueError("只有 Word（.doc/.docx）和 Markdown 文件可以转换为可搜索PDF")


def convert_and_import(database, document_row) -> tuple[str, int, str]:
    """Convert an imported document and add the derived PDF to the same category."""
    source = Path(document_row["stored_path"])
    output_name = f"{Path(document_row['name']).stem}（可搜索）.pdf"
    with database.connect() as connection:
        existing = connection.execute(
            """SELECT stored_path FROM documents
               WHERE name = ? COLLATE NOCASE AND category = ? COLLATE NOCASE LIMIT 1""",
            (output_name, str(document_row["category"])),
        ).fetchone()
    if existing and Path(existing["stored_path"]).exists():
        return "skipped", 0, output_name
    with tempfile.TemporaryDirectory() as temporary:
        output = Path(temporary) / output_name
        convert_document_to_searchable_pdf(source, output)
        status, chunks = database.import_file(
            output, duplicate_action="skip", category=str(document_row["category"])
        )
    return status, chunks, output_name
