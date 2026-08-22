#!/usr/bin/env python3
"""Convert Manin DjVu scan to FB2 via Tesseract OCR."""

import html
import re
import subprocess
import tempfile
from pathlib import Path

DJVU = Path(
    "/home/ubuntu/.cursor/projects/workspace/uploads/"
    "_______._._-_________________________________________-_1980_235b.djvu"
)
OUTPUT = Path("/workspace/Convert/Manin_Vychislimoe_i_nevychislimoe_1980.fb2")
PAGE_OFFSET = 2  # book_page = djvu_page - PAGE_OFFSET

STRUCTURE = [
    {"title": "Предисловие", "page": 3},
    {"title": "Введение", "page": 5},
    {
        "title": "Глава I. Рекурсивные функции и алгоритмы",
        "page": 16,
        "sections": [
            {"title": "1. Интуитивная вычислимость", "page": 16},
            {"title": "2. Частично рекурсивные функции", "page": 20},
            {"title": "3. Образцы рекурсивности", "page": 25},
            {"title": "4. Перечислимые и разрешимые множества", "page": 29},
            {"title": "5. Элементы рекурсивной геометрии", "page": 39},
            {"title": "6. Конструктивные объекты и алгоритмы", "page": 44},
        ],
    },
    {
        "title": "Глава II. Диофантовы множества и алгоритмическая неразрешимость",
        "page": 46,
        "sections": [
            {"title": "1. Основные результаты", "page": 46},
            {"title": "2. План доказательства", "page": 49},
            {"title": "3. Перечислимые множества являются D-множествами", "page": 51},
            {"title": "4. Редукция", "page": 53},
            {"title": "5. Конструкция специального диофантова множества", "page": 56},
            {"title": "6. График экспоненты диофантов", "page": 61},
            {"title": "7. Графики факториала и биномиальных коэффициентов диофантовы", "page": 62},
            {"title": "8. Дополнения", "page": 64},
        ],
    },
    {
        "title": "Глава III. Сложность и случайность",
        "page": 65,
        "sections": [
            {"title": "1. Версальные семейства", "page": 65},
            {"title": "2. Сложность по Колмогорову", "page": 68},
            {"title": "3. Сложность и случайность", "page": 73},
        ],
    },
    {
        "title": "Глава IV. Формальные языки и вычислимость",
        "page": 74,
        "sections": [
            {"title": "1. Арифметика синтаксиса", "page": 74},
            {"title": "2. Синтаксический анализ", "page": 80},
            {"title": "3. Перечислимость выводимых формул", "page": 85},
        ],
    },
    {
        "title": "Глава V. Теорема Гёделя",
        "page": 87,
        "sections": [
            {"title": "1. Принцип неполноты", "page": 87},
            {"title": "2. Неперечислимость истинных формул", "page": 88},
            {"title": "3. О длине доказательств", "page": 91},
            {"title": "4. Арифметическая иерархия", "page": 93},
            {"title": "5. Продуктивность арифметической истины", "page": 96},
            {"title": "6. Вычислимые функции с очень быстрым ростом", "page": 99},
        ],
    },
    {
        "title": "Глава VI. Рекурсивные группы",
        "page": 101,
        "sections": [
            {"title": "1. Основной результат и его следствия", "page": 101},
            {"title": "2. Свободные произведения и HNN-расширения", "page": 104},
            {"title": "3. Вложения в группы с двумя образующими", "page": 107},
            {"title": "4. Хорошие подгруппы", "page": 108},
            {"title": "5. Ограниченные системы образующих", "page": 111},
            {"title": "6. Окончание доказательства", "page": 116},
        ],
    },
    {"title": "Список литературы", "page": 123},
]


def book_to_djvu(book_page: int) -> int:
    return book_page + PAGE_OFFSET


def ocr_page(djvu_page: int, tmpdir: Path) -> str:
    ppm = tmpdir / f"p{djvu_page:04d}.ppm"
    png = tmpdir / f"p{djvu_page:04d}.png"
    subprocess.run(
        ["ddjvu", "-format=ppm", f"-page={djvu_page}", str(DJVU), str(ppm)],
        check=True,
        capture_output=True,
    )
    subprocess.run(
        [
            "convert",
            str(ppm),
            "-resize",
            "2400x",
            "-colorspace",
            "Gray",
            "-normalize",
            str(png),
        ],
        check=True,
        capture_output=True,
    )
    ppm.unlink(missing_ok=True)
    result = subprocess.run(
        ["tesseract", str(png), "stdout", "-l", "rus", "--psm", "1"],
        capture_output=True,
        text=True,
        check=True,
    )
    png.unlink(missing_ok=True)
    return clean_ocr_text(result.stdout)


def clean_ocr_text(text: str) -> str:
    lines = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            lines.append("")
            continue
        # Drop isolated page numbers.
        if re.fullmatch(r"\d{1,3}", line):
            continue
        # Drop figure captions at line start when alone.
        if re.fullmatch(r"Рис\.\s*\d+", line):
            continue
        lines.append(line)
    text = "\n".join(lines)
    text = re.sub(r"\n{3,}", "\n\n", text)
    # Fix common OCR artifacts.
    text = text.replace("—", "—")
    text = re.sub(r"(\w)-\n(\w)", r"\1\2", text)
    return text.strip()


def text_to_paragraphs(text: str) -> list[str]:
    if not text:
        return []
    paragraphs = []
    for block in re.split(r"\n\s*\n", text):
        block = " ".join(block.split())
        if block:
            paragraphs.append(block)
    return paragraphs


def esc(text: str) -> str:
    return html.escape(text, quote=False)


def flatten_sections() -> list[tuple[str, int, int | None]]:
    """Return (title, start_book_page, end_book_page_or_None)."""
    items: list[tuple[str, int]] = []

    def walk(node: dict) -> None:
        if "sections" in node:
            items.append((node["title"], node["page"]))
            for sec in node["sections"]:
                items.append((sec["title"], sec["page"]))
        else:
            items.append((node["title"], node["page"]))

    for node in STRUCTURE:
        walk(node)

    ranges: list[tuple[str, int, int | None]] = []
    for i, (title, start) in enumerate(items):
        end = items[i + 1][1] - 1 if i + 1 < len(items) else 122
        ranges.append((title, start, end))
    return ranges


def render_section(title: str, paragraphs: list[str], level: int = 1) -> str:
    tag = "section"
    parts = [f"<{tag}>"]
    parts.append("<title>")
    parts.append(f"<p>{esc(title)}</p>")
    parts.append("</title>")
    for p in paragraphs:
        parts.append(f"<p>{esc(p)}</p>")
    parts.append(f"</{tag}>")
    return "\n".join(parts)


def build_body(page_text: dict[int, str]) -> str:
    sections = flatten_sections()
    body_parts = ["<body>", "<title>", "<p>Вычислимое и невычислимое</p>", "</title>"]

    chapter_buf: list[str] = []
    chapter_title: str | None = None

    def flush_chapter() -> None:
        nonlocal chapter_buf, chapter_title
        if chapter_title and chapter_buf:
            chapter_buf.append("</section>")
            body_parts.append("\n".join(chapter_buf))
        chapter_buf = []
        chapter_title = None

    for title, start, end in sections:
        chunks: list[str] = []
        for book_page in range(start, (end or start) + 1):
            chunks.append(page_text.get(book_page, ""))
        paragraphs = text_to_paragraphs("\n\n".join(chunks))

        is_chapter = title.startswith("Глава ")

        if is_chapter:
            flush_chapter()
            chapter_title = title
            chapter_buf = ["<section>", "<title>", f"<p>{esc(title)}</p>", "</title>"]
            continue

        if chapter_buf is not None and chapter_title is not None:
            chapter_buf.append(render_section(title, paragraphs))
        else:
            body_parts.append(render_section(title, paragraphs))

    flush_chapter()
    body_parts.append("</body>")
    return "\n".join(body_parts)


def main() -> None:
    result = subprocess.run(
        ["djvused", str(DJVU), "-e", "n"],
        capture_output=True,
        text=True,
        check=True,
    )
    total_pages = int(result.stdout.strip())
    print(f"DjVu pages: {total_pages}")

    page_text: dict[int, str] = {}
    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        for djvu_page in range(1, total_pages + 1):
            book_page = djvu_page - PAGE_OFFSET
            if book_page < 1 or book_page > 122:
                continue
            print(f"OCR DjVu {djvu_page} -> book page {book_page}...", flush=True)
            page_text[book_page] = ocr_page(djvu_page, tmpdir)

    body = build_body(page_text)

    fb2 = f"""<?xml version="1.0" encoding="UTF-8"?>
<FictionBook xmlns="http://www.gribuser.ru/xml/fictionbook/2.0" xmlns:l="http://www.w3.org/1999/xlink">
<description>
  <title-info>
    <genre>science_math</genre>
    <author>
      <first-name>Юрий</first-name>
      <middle-name>Иванович</middle-name>
      <last-name>Манин</last-name>
    </author>
    <book-title>Вычислимое и невычислимое</book-title>
    <annotation>
      <p>Книга посвящена доказательству существования невычислимых функций и алгоритмически неразрешимых задач. Обсуждаются проблемы оценки сложности вычислений и алгоритмов.</p>
    </annotation>
    <keywords>вычислимость, рекурсивные функции, теория алгоритмов, кибернетика</keywords>
    <date value="1980">1980</date>
    <lang>ru</lang>
    <src-lang>ru</src-lang>
    <sequence name="Кибернетика"/>
    <publisher>
      <first-name></first-name>
      <last-name>Советское радио</last-name>
    </publisher>
  </title-info>
  <document-info>
    <author>
      <nickname>OCR conversion</nickname>
    </author>
    <program-used>djvulibre + Tesseract OCR</program-used>
    <date value="2026-08-22">2026-08-22</date>
    <src-url></src-url>
    <id>manin-vychislimoe-1980</id>
    <version>1.0</version>
  </document-info>
  <publish-info>
    <book-name>Вычислимое и невычислимое</book-name>
    <publisher>Советское радио</publisher>
    <city>Москва</city>
    <year>1980</year>
  </publish-info>
</description>
{body}
</FictionBook>
"""

    OUTPUT.write_text(fb2, encoding="utf-8")
    print(f"Written: {OUTPUT} ({OUTPUT.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
