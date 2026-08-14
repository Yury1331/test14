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
    run = sub.add_run("№ КП-2026-08-14-001, редакция 2 от 14 августа 2026 г.")
    set_run_font(run, size=13, bold=True, color=NAVY, name="Calibri")

    add_body(
        doc,
        "Предмет: поставка 2 видеорегистраторов TRASSIR NeuroStation Astra 9800R/128-S и 51 камеры "
        "(1× НИЦ 1143, 49× РУВЕР 4MBIR-28-TMLW, 1× PTZ 2BPBDD-4794-20). Все камеры — со СТ-1. "
        "Валюта — российский рубль. НДС включён в цены (ставка 22%). Срок действия цен — до 28.08.2026.",
        size=11,
    )
    add_body(
        doc,
        "Документ — ориентировочный расчёт по открытым ценам на 14.08.2026. Итоговая цена, наличие, "
        "срок и выдача СТ-1 на партию фиксируются в официальном КП DSSL, НИЦ «Технологии» и ООО «Айпи Плюс».",
        size=10,
        italic=True,
        color=MUTED,
    )

    meta = [
        ("Заказчик", "______________________________"),
        ("ИНН / КПП", "______________________________"),
        ("Поставщик", "к заполнению (DSSL / TRASSIR, НИЦ «Технологии», ООО «Айпи Плюс» / РУВЕР)"),
        ("Срок поставки", "регистраторы: 5–15 раб. дней при наличии; камеры — под заказ"),
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
        ["№", "Наименование", "Артикул", "СТ-1", "Ед.", "Кол-во", "Цена с НДС, ₽", "Сумма с НДС, ₽"],
        [
            [
                "1",
                "IP-видеорегистратор TRASSIR NeuroStation Astra 9800R/128-S, 128 кан., Astra Linux SE «Смоленск», 8×HDD (диски не входят), 2× Neuro Detector",
                "NeuroStation 9800R/128-S",
                "да",
                "шт.",
                "2",
                "477 740",
                "955 480",
            ],
            [
                "2",
                "Купольная IP 4 Мп, мотор 2.7–13.5 мм, ИК 35 м, IP66/IK10, ПО Nexus. НИЦ модель 1143",
                "4MP-DOM-2.7-13.5M Nexus",
                "да",
                "шт.",
                "1",
                "23 800",
                "23 800",
            ],
            [
                "3",
                "Уличная цилиндрическая IP 4 Мп Mini Bullet, 2.8 мм, ИК + белый свет, микрофон. Производитель РУВЕР",
                "4MBIR-28-TMLW",
                "да",
                "шт.",
                "49",
                "18 500",
                "906 500",
            ],
            [
                "4",
                "PTZ уличная 2 Мп, зум ×20 (4.7–94 мм), IP66. Самая недорогая PTZ РУВЕР с СТ-1 в серии 2 Мп",
                "2BPBDD-4794-20",
                "да",
                "шт.",
                "1",
                "92 750",
                "92 750",
            ],
        ],
        money_cols={6, 7},
    )

    add_body(doc, "Итого по запросу (позиции 1–4). Камеры: 1+49+1 = 51 шт.", size=11)
    add_table(
        doc,
        ["Показатель", "Сумма, ₽"],
        [
            ["Всего с НДС", "1 978 530"],
            ["в том числе НДС 22%", "356 784,10"],
            ["без НДС", "1 621 745,90"],
        ],
        money_cols={1},
    )

    add_heading_bar(doc, "2. Что нужно докупить, чтобы система заработала")
    add_body(
        doc,
        "Регистратор не содержит лицензий на IP-камеры и HDD. Камеры НИЦ/РУВЕР — не native TRASSIR, "
        "на каждый канал нужна лицензия AnyIP.",
    )
    add_table(
        doc,
        ["№", "Наименование", "Артикул / код", "Ед.", "Кол-во", "Цена с НДС, ₽", "Сумма с НДС, ₽"],
        [
            [
                "5",
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
    run = p.add_run("Итого запрос + лицензии AnyIP: 2 238 120 ₽")
    set_run_font(run, size=12, bold=True, color=NAVY, name="Calibri")
    run2 = p.add_run("  (в т.ч. НДС 22% — 403 595,41 ₽).")
    set_run_font(run2, size=11, color=BLACK)

    add_heading_bar(doc, "3. Рекомендуемые опции (не в исходном запросе)")
    add_body(
        doc,
        "Оценка архива: 51 камера, в основном 4 Мп, H.265, средний битрейт 3 Мбит/с → ≈ 1,65 ТБ/сутки. "
        "На 30 суток нужно ≈ 50 ТБ полезного объёма.",
    )
    add_table(
        doc,
        ["№", "Наименование", "Кол-во", "Цена с НДС, ₽", "Сумма с НДС, ₽", "Назначение"],
        [
            [
                "6",
                "HDD 10 ТБ Seagate SkyHawk AI ST10000VE001 (рекомендован DSSL для NeuroStation 8-disk)",
                "8",
                "55 200",
                "441 600",
                "≈30 суток архива на одном NVR",
            ],
            [
                "7",
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
            ["Запрос (п. 1–4)", "1 978 530"],
            ["Работоспособный минимум (п. 1–5)", "2 238 120"],
            ["+ архив 30 суток (п. 1–6)", "2 679 720"],
            ["+ аналитика на все камеры (п. 1–7)", "2 989 920"],
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

    add_heading_bar(doc, "4. Кратко по оборудованию и СТ-1")
    p = doc.add_paragraph()
    run = p.add_run("4.1. TRASSIR NeuroStation Astra 9800R/128-S — 2 шт. ")
    set_run_font(run, size=11, bold=True)
    run = p.add_run(
        "128 IP-каналов, полоса 790 Мбит/с, 8 отсеков HDD, стойка 2U, Astra Linux SE «Смоленск», "
        "реестры Минпромторга №719 / №878, СТ-1. Цена 477 740 ₽ — открытая розница; у DSSL проектная."
    )
    set_run_font(run, size=11)

    p = doc.add_paragraph()
    run = p.add_run("4.2. НИЦ модель 1143, 4MP-DOM-2.7-13.5M Nexus — 1 шт. ")
    set_run_font(run, size=11, bold=True)
    run = p.add_run(
        "Купол 4 Мп, Sony IMX335, мотор 2.7–13.5 мм, ИК до 35 м, IP66/IK10, PoE, −45…+50 °C, ПО Nexus "
        "(реестр Минцифры №12466). Изготовитель ООО «НИЦ Технологии», Россия. СТ-1 на партию — у завода. "
        "Спецификация: nic-tech.ru, модель 1143."
    )
    set_run_font(run, size=11)

    p = doc.add_paragraph()
    run = p.add_run("4.3. РУВЕР 4MBIR-28-TMLW — 49 шт. ")
    set_run_font(run, size=11, bold=True)
    run = p.add_run(
        "Mini Bullet 4 Мп, фикс 2.8 мм, ИК + белый свет, микрофон. Производитель RUWARE, продавец "
        "ООО «Айпи Плюс». В каталоге ipplus.ru флаг СТ-1. "
        "https://www.ipplus.ru/catalog/ip-kamery/ip-kamery-4mp/4mbir-28-tmlw"
    )
    set_run_font(run, size=11)

    p = doc.add_paragraph()
    run = p.add_run("4.4. PTZ РУВЕР 2BPBDD-4794-20 — 1 шт. ")
    set_run_font(run, size=11, bold=True)
    run = p.add_run(
        "Самая недорогая поворотная камера того же производителя (РУВЕР / Айпи Плюс) с СТ-1: 2 Мп, "
        "зум ×20, 4.7–94 мм, IP66. Модели ×10 в каталоге без СТ-1 не предлагаются. "
        "https://www.ipplus.ru/catalog/ptz-kamery/ptz-kamery-2mp/2bpbdd-4794-20"
    )
    set_run_font(run, size=11)

    add_heading_bar(doc, "5. Условия (типовые, к согласованию)")
    terms = [
        "Цены в рублях с НДС 22%, ориентир открытого рынка на 14.08.2026, редакция 2.",
        "В поставке камер — сертификат СТ-1 на партию и/или номер реестровой записи РЭП (ПП РФ №878).",
        "Оплата: 100% предоплата или 70/30 — по договору.",
        "Поставка: самовывоз со склада дилера в РФ либо доставка за счёт заказчика.",
        "Регистраторы поставляются без HDD и без лицензий AnyIP.",
        "Цены камер РУВЕР фиксируются КП ООО «Айпи Плюс»; НИЦ 1143 — КП НИЦ / дилера.",
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
            ["НИЦ «Технологии» (камера 1143)", "nic-tech.ru, support@nic-tech.ru, 8 (800) 555-47-65"],
            ["ООО «Айпи Плюс» / РУВЕР (4MBIR, PTZ)", "8 804 333-73-02, info@ipplus.ru, ipplus.ru"],
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
