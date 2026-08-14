#!/usr/bin/env python3
"""Сборка КП № КП-2026-08-14-001 в формате Microsoft Word (.docx)."""

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

NAVY = RGBColor(0x1F, 0x4E, 0x79)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)
MUTED = RGBColor(0x59, 0x59, 0x59)
RED = RGBColor(0xC0, 0x00, 0x00)

OUT = (
    "/workspace/kp/2026-08-14-neurostation-4bhnn0/"
    "КП-2026-08-14-001_NeuroStation_Astra_9800R_4BHNN0.docx"
)


def set_run_font(run, *, size=11, bold=False, color=BLACK, name="Times New Roman"):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    r = run._element
    rpr = r.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), name)
    rfonts.set(qn("w:hAnsi"), name)
    rfonts.set(qn("w:cs"), name)
    rfonts.set(qn("w:eastAsia"), name)


def shade_cell(cell, hex_color: str) -> None:
    tc = cell._tc
    tcpr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcpr.append(shd)


def set_cell_border(cell) -> None:
    tc = cell._tc
    tcpr = tc.get_or_add_tcPr()
    tc_borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "BFBFBF")
        tc_borders.append(el)
    tcpr.append(tc_borders)


def set_cell_text(cell, text, *, bold=False, size=9, align="left", color=BLACK, fill=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }[align]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(str(text))
    set_run_font(run, size=size, bold=bold, color=color)
    if fill:
        shade_cell(cell, fill)
    set_cell_border(cell)
    cell.vertical_alignment = 1  # center


def add_heading_bar(doc, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    set_run_font(run, size=13, bold=True, color=NAVY, name="Calibri")


def add_body(doc, text: str, *, size=11, italic=False, color=BLACK, space_after=6) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    run = p.add_run(text)
    set_run_font(run, size=size, color=color)
    run.italic = italic


def add_table(doc, headers, rows, col_widths=None, money_cols=None):
    money_cols = money_cols or set()
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for i, h in enumerate(headers):
        set_cell_text(table.rows[0].cells[i], h, bold=True, size=8, align="center", color=WHITE, fill="1F4E79")
    for r_idx, row in enumerate(rows):
        fill = "F2F2F2" if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, val in enumerate(row):
            align = "right" if c_idx in money_cols else ("center" if c_idx in (0, 3, 4) else "left")
            set_cell_text(
                table.rows[r_idx + 1].cells[c_idx],
                val,
                size=8,
                align=align,
                fill=fill,
                bold=c_idx in money_cols and r_idx == len(rows) - 1 and False,
            )
    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Cm(w)
    return table


def set_narrow_margins(doc) -> None:
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)


def add_footer(doc) -> None:
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("КП-2026-08-14-001  ·  ориентировочный расчёт  ·  цены с НДС 22%  ·  стр. ")
    set_run_font(run, size=8, color=MUTED, name="Calibri")

    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")
    run2 = p.add_run()
    run2._r.append(fld_char_begin)
    run2._r.append(instr)
    run2._r.append(fld_char_end)
    set_run_font(run2, size=8, color=MUTED, name="Calibri")


def build() -> None:
    doc = Document()
    set_narrow_margins(doc)
    add_footer(doc)

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(2)
    run = title.add_run("КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ")
    set_run_font(run, size=18, bold=True, color=NAVY, name="Calibri")

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.paragraph_format.space_after = Pt(10)
    run = sub.add_run("№ КП-2026-08-14-001 от 14 августа 2026 г.")
    set_run_font(run, size=13, bold=True, color=NAVY, name="Calibri")

    add_body(
        doc,
        "Предмет: поставка видеорегистраторов TRASSIR NeuroStation Astra 9800R/128-S "
        "и цифровых камер видеонаблюдения 4BHNN0-0-0-0. Валюта — российский рубль. "
        "НДС включён в цены (ставка 22%). Срок действия цен — 14 календарных дней (до 28.08.2026).",
        size=11,
    )
    add_body(
        doc,
        "Документ — ориентировочный расчёт поставки по открытым ценам на 14.08.2026. "
        "Итоговая цена, наличие и срок фиксируются в официальном КП DSSL / авторизованного "
        "дилера TRASSIR и НИЦ «Технологии».",
        size=10,
        italic=True,
        color=MUTED,
    )

    meta = [
        ("Заказчик", "______________________________"),
        ("ИНН / КПП", "______________________________"),
        ("Поставщик", "к заполнению (ориентир: дилеры TRASSIR / DSSL, НИЦ «Технологии»)"),
        ("Срок поставки", "регистраторы: 5–15 раб. дней при наличии; камеры НИЦ: под заказ"),
        ("Оплата", "по договору (типовой ориентир: 100% предоплата или 70/30)"),
    ]
    table = doc.add_table(rows=len(meta), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (k, v) in enumerate(meta):
        set_cell_text(table.rows[i].cells[0], k, bold=True, size=10, fill="D6E4F0")
        set_cell_text(table.rows[i].cells[1], v, size=10)
        table.rows[i].cells[0].width = Cm(4.5)
        table.rows[i].cells[1].width = Cm(13.5)

    add_heading_bar(doc, "1. Спецификация по запросу")
    add_table(
        doc,
        ["№", "Наименование", "Артикул", "Ед.", "Кол-во", "Цена с НДС, ₽", "Сумма с НДС, ₽"],
        [
            [
                "1",
                "IP-видеорегистратор TRASSIR NeuroStation Astra 9800R/128-S, 128 кан., Astra Linux SE «Смоленск», 8×HDD 3.5\" (диски не входят), 2 лицензии Neuro Detector",
                "NeuroStation 9800R/128-S",
                "шт.",
                "2",
                "477 740",
                "955 480",
            ],
            [
                "2",
                "Камера видеонаблюдения цифровая 4 Мп, прошивка Nexus. Артикул подтвердить у НИЦ (корпус/объектив в коде не заданы)",
                "4BHNN0-0-0-0",
                "шт.",
                "51",
                "22 000",
                "1 122 000",
            ],
        ],
        money_cols={5, 6},
    )

    add_body(doc, "Итого по запросу (позиции 1–2)", size=11)
    add_table(
        doc,
        ["Показатель", "Сумма, ₽"],
        [
            ["Всего с НДС", "2 077 480"],
            ["в том числе НДС 22%", "374 627,54"],
            ["без НДС", "1 702 852,46"],
        ],
        money_cols={1},
    )

    add_heading_bar(doc, "2. Что нужно докупить, чтобы система заработала")
    add_body(
        doc,
        "Регистратор не содержит лицензий на IP-камеры и HDD. Камеры 4BHNN0 — не native TRASSIR, "
        "на каждый канал нужна лицензия AnyIP.",
    )
    add_table(
        doc,
        ["№", "Наименование", "Артикул / код", "Ед.", "Кол-во", "Цена с НДС, ₽", "Сумма с НДС, ₽"],
        [
            [
                "3",
                "Лицензия TRASSIR AnyIP (TRASSIR OS) — подключение 1 IP-камеры",
                "DSSL 57885",
                "шт.",
                "51",
                "5 090",
                "259 590",
            ],
        ],
        money_cols={5, 6},
    )
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    run = p.add_run("Итого запрос + лицензии AnyIP: 2 337 070 ₽")
    set_run_font(run, size=12, bold=True, color=NAVY, name="Calibri")
    run2 = p.add_run("  (в т.ч. НДС 22% — 421 438,85 ₽).")
    set_run_font(run2, size=11, color=BLACK)

    add_heading_bar(doc, "3. Рекомендуемые опции (не в исходном запросе)")
    add_body(
        doc,
        "Оценка архива: 51 камера × 4 Мп, H.265, средний битрейт 3 Мбит/с → ≈ 1,65 ТБ/сутки "
        "непрерывной записи. На 30 суток нужно ≈ 50 ТБ полезного объёма.",
    )
    add_table(
        doc,
        ["№", "Наименование", "Кол-во", "Цена с НДС, ₽", "Сумма с НДС, ₽", "Назначение"],
        [
            [
                "4",
                "HDD 10 ТБ Seagate SkyHawk AI ST10000VE001 (рекомендован DSSL для NeuroStation 8-disk)",
                "8",
                "55 200",
                "441 600",
                "≈30 суток архива на одном NVR",
            ],
            [
                "5",
                "TRASSIR Neuro Detector, доп. каналы (в 2 NVR уже 4 лицензии)",
                "47",
                "6 600",
                "310 200",
                "аналитика на все 51 камеру",
            ],
        ],
        money_cols={3, 4},
    )

    add_body(doc, "Сводка комплектов", size=11)
    add_table(
        doc,
        ["Комплект", "Сумма с НДС, ₽"],
        [
            ["Запрос (п. 1–2)", "2 077 480"],
            ["Работоспособный минимум (п. 1–3)", "2 337 070"],
            ["+ архив 30 суток (п. 1–4)", "2 778 670"],
            ["+ аналитика на все камеры (п. 1–5)", "3 088 870"],
        ],
        money_cols={1},
    )
    add_body(
        doc,
        "В расчёт не входят: PoE-коммутаторы, СКС, шкаф 19\", ИБП, мониторы, пуско-наладка, "
        "доставка (если не самовывоз).",
        size=10,
        italic=True,
        color=MUTED,
    )

    add_heading_bar(doc, "4. Кратко по оборудованию")
    p = doc.add_paragraph()
    run = p.add_run("4.1. TRASSIR NeuroStation Astra 9800R/128-S — 2 шт. ")
    set_run_font(run, size=11, bold=True)
    run = p.add_run(
        "128 IP-каналов, полоса 790 Мбит/с, 8 отсеков HDD, стойка 2U, Astra Linux SE «Смоленск», "
        "реестры Минпромторга №719 / №878, СТ-1. Два регистратора на 51 камеру дают запас по каналам "
        "(256 vs 51): резерв, два объекта или рост системы. Запись можно вести на одном NVR, второй "
        "держать как горячий резерв. Цена 477 740 ₽ — открытая розница trasrussia.ru; у DSSL цена проектная."
    )
    set_run_font(run, size=11)

    p = doc.add_paragraph()
    run = p.add_run("4.2. Камера 4BHNN0-0-0-0 — 51 шт. ")
    set_run_font(run, size=11, bold=True)
    run = p.add_run(
        "По коду — 4 Мп, прошивка Nexus (NN), базовая комплектация 0-0-0-0; производитель с высокой "
        "вероятностью НИЦ «Технологии». Корпус и объектив в артикуле не заданы — до заказа подтвердить: "
        "цилиндр / купол / корпусная, фикс или мотор, ИК, PoE. Цена 22 000 ₽ — ориентир опта 15–25 тыс. ₽."
    )
    set_run_font(run, size=11)

    add_heading_bar(doc, "5. Условия (типовые, к согласованию)")
    terms = [
        "Цены в рублях с НДС 22%, ориентир открытого рынка на 14.08.2026.",
        "Оплата: 100% предоплата или 70/30 — по договору.",
        "Поставка: самовывоз со склада дилера в РФ либо доставка за счёт заказчика (тариф отдельно).",
        "Гарантия: TRASSIR — по паспорту производителя (типично 3–5 лет на NVR); камеры НИЦ — уточняется (часто 3 года).",
        "Регистраторы поставляются без HDD и без лицензий AnyIP.",
        "Цена камеры 4BHNN0-0-0-0 — ориентир; фиксируется после подтверждения артикула и КП НИЦ / дилера.",
        "DSSL даёт проектную цену на NeuroStation — партнёрская может быть ниже 477 740 ₽.",
    ]
    for i, t in enumerate(terms, 1):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Cm(0.5)
        run = p.add_run(f"{i}. {t}")
        set_run_font(run, size=11)

    add_heading_bar(doc, "6. Контакты для официального КП")
    add_table(
        doc,
        ["Поставщик", "Контакт"],
        [
            ["DSSL / TRASSIR", "8 (800) 100-91-12, dssl.ru, код товара NVR 80222"],
            ["НИЦ «Технологии»", "nic-tech.ru, support@nic-tech.ru, 8 (800) 555-47-65"],
        ],
    )

    sig = doc.add_paragraph()
    sig.paragraph_format.space_before = Pt(24)
    run = sig.add_run("Поставщик / исполнитель                          Заказчик")
    set_run_font(run, size=11, bold=True)

    sig2 = doc.add_paragraph()
    sig2.paragraph_format.space_before = Pt(18)
    run = sig2.add_run("________________ / ________________              ________________ / ________________")
    set_run_font(run, size=11)
    sig3 = doc.add_paragraph()
    run = sig3.add_run("подпись, ФИО, дата                                         подпись, ФИО, дата")
    set_run_font(run, size=9, color=MUTED)

    doc.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
