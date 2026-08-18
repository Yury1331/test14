#!/usr/bin/env python3
"""Сборка КП аптеки: XLSX, DOCX, PDF."""

from __future__ import annotations

from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parent
TODAY = date(2026, 8, 18)
KP_NO = "КП-Аптека-2026-08-18"

ITEMS = [
    {
        "name": "IP-камера Hikvision DS-2CD2047G2-LU(C)(2.8mm), 4 Мп ColorVu, 2.8 мм, LED до 40 м, микрофон, IP67",
        "sku": "DS-2CD2047G2-LU(C)(2.8mm)",
        "unit": "шт.",
        "qty": 16,
        "price": 12990,
    },
    {
        "name": "Монтажная коробка Hikvision DS-1280ZJ-S (цилиндр 20-линии)",
        "sku": "DS-1280ZJ-S",
        "unit": "шт.",
        "qty": 16,
        "price": 1290,
    },
    {
        "name": "IP-видеорегистратор Hikvision DS-7616NXI-K2, 16 каналов, 2×SATA до 16 ТБ, без PoE",
        "sku": "DS-7616NXI-K2",
        "unit": "шт.",
        "qty": 1,
        "price": 22490,
    },
    {
        "name": "Жёсткий диск Seagate SkyHawk 12 ТБ (ST12000VE001 / актуальный аналог SkyHawk)",
        "sku": "ST12000VE001",
        "unit": "шт.",
        "qty": 2,
        "price": 28900,
    },
    {
        "name": "PoE-коммутатор Hikvision DS-3E1326P-EI, 24×FE PoE + 2×GE combo, бюджет 370 Вт (HiPoE)",
        "sku": "DS-3E1326P-EI",
        "unit": "шт.",
        "qty": 1,
        "price": 32990,
    },
    {
        "name": 'Шкаф телекоммуникационный настенный 12U 19" 600×600, дверь стекло',
        "sku": "12U-600x600",
        "unit": "шт.",
        "qty": 1,
        "price": 16900,
    },
    {
        "name": "ИБП 2000 ВА / ≥1200 Вт, чистая синусоида (Ippon Smart Winner II 2000 или аналог)",
        "sku": "UPS-2000-SINE",
        "unit": "шт.",
        "qty": 1,
        "price": 28900,
    },
    {
        "name": 'Комплект размещения в шкафу: полка 19" 2U + PDU 8 розеток 1U',
        "sku": "SHELF-PDU",
        "unit": "компл.",
        "qty": 1,
        "price": 4900,
    },
]

for row in ITEMS:
    row["sum"] = row["qty"] * row["price"]

TOTAL = sum(r["sum"] for r in ITEMS)
NET = round(TOTAL / 1.22, 2)
VAT = round(TOTAL - NET, 2)

NAVY = "1F4E79"
NAVY_RGB = RGBColor(0x1F, 0x4E, 0x79)


def money(n: float) -> str:
    return f"{n:,.2f}".replace(",", " ").replace(".", ",") if isinstance(n, float) and not n.is_integer() else f"{int(n):,}".replace(",", " ")


def money_ru(n: float) -> str:
    if abs(n - round(n)) < 0.001:
        return f"{int(round(n)):,}".replace(",", " ")
    s = f"{n:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", " ")


def set_run_font(run, name="Calibri", size=11, bold=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = color
    r = run._element.get_or_add_rPr()
    rFonts = r.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        r.append(rFonts)
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:eastAsia"), name)
    rFonts.set(qn("w:cs"), name)


def shade_cell(cell, fill_hex: str):
    tc = cell._tePr if hasattr(cell, "_tePr") else cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill_hex)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), "8FAADC")
        tcBorders.append(el)
    tcPr.append(tcBorders)


def build_xlsx() -> Path:
    wb = Workbook()

    # --- КП ---
    ws = wb.active
    ws.title = "КП"
    thin = Border(
        left=Side(style="thin", color="8FAADC"),
        right=Side(style="thin", color="8FAADC"),
        top=Side(style="thin", color="8FAADC"),
        bottom=Side(style="thin", color="8FAADC"),
    )
    fill_navy = PatternFill("solid", fgColor=NAVY)
    fill_head = PatternFill("solid", fgColor="2E75B6")
    fill_total = PatternFill("solid", fgColor="FFF2CC")
    fill_alt = PatternFill("solid", fgColor="D6EAF8")
    font_white = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    font_title = Font(name="Calibri", bold=True, color="FFFFFF", size=16)
    font_bold = Font(name="Calibri", bold=True, size=11)
    font_n = Font(name="Calibri", size=11)
    wrap = Alignment(wrap_text=True, vertical="center")
    center = Alignment(wrap_text=True, vertical="center", horizontal="center")
    right = Alignment(wrap_text=True, vertical="center", horizontal="right")

    ws.merge_cells("A1:G1")
    ws["A1"] = "ООО «Эффорт Юг»  ·  коммерческое предложение"
    ws["A1"].font = font_title
    ws["A1"].fill = fill_navy
    ws["A1"].alignment = Alignment(vertical="center", horizontal="left")
    ws.row_dimensions[1].height = 28

    meta = [
        ("Поставщик", "ООО «Эффорт Юг»"),
        ("ИНН / ОГРН", "6165231466 / 1226100001148"),
        ("Адрес", "344064, г. Ростов-на-Дону, ул. Вавилова, д. 71 В стр. 2, этаж 2 помещ. 2 ж"),
        ("Документ", f"{KP_NO} от {TODAY.strftime('%d.%m.%Y')}"),
        ("Объект", "Аптека, система видеонаблюдения 16 камер ColorVu"),
        ("Основание", "Корректировка исходного КП (8 каналов / 5 камер AcuSense / 2 ТБ)"),
    ]
    r = 3
    for k, v in meta:
        ws[f"A{r}"] = k
        ws[f"A{r}"].font = font_bold
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws[f"B{r}"] = v
        ws[f"B{r}"].font = font_n
        r += 1

    header_row = r + 1
    headers = ["№", "Наименование", "Артикул", "Ед.", "Кол-во", "Цена с НДС, ₽", "Сумма с НДС, ₽"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(header_row, col, h)
        cell.font = font_white
        cell.fill = fill_head
        cell.alignment = center
        cell.border = thin

    for i, item in enumerate(ITEMS, 1):
        row = header_row + i
        values = [i, item["name"], item["sku"], item["unit"], item["qty"], item["price"], item["sum"]]
        for col, val in enumerate(values, 1):
            cell = ws.cell(row, col, val)
            cell.font = font_n
            cell.border = thin
            cell.alignment = wrap if col == 2 else center
            if col in (6, 7):
                cell.number_format = '#,##0'
                cell.alignment = right
            if i % 2 == 0:
                cell.fill = fill_alt

    total_row = header_row + len(ITEMS) + 1
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=6)
    ws.cell(total_row, 1, "Итого общая стоимость оборудования с НДС 22%").font = font_bold
    ws.cell(total_row, 1).fill = fill_total
    ws.cell(total_row, 1).alignment = Alignment(vertical="center", horizontal="right")
    ws.cell(total_row, 7, TOTAL).font = font_bold
    ws.cell(total_row, 7).fill = fill_total
    ws.cell(total_row, 7).number_format = '#,##0'
    for col in range(1, 8):
        ws.cell(total_row, col).fill = fill_total
        ws.cell(total_row, col).border = thin

    vat_row = total_row + 1
    ws.merge_cells(start_row=vat_row, start_column=1, end_row=vat_row, end_column=6)
    ws.cell(vat_row, 1, "в том числе НДС 22%").alignment = Alignment(horizontal="right")
    ws.cell(vat_row, 7, VAT).number_format = '#,##0.00'
    net_row = vat_row + 1
    ws.merge_cells(start_row=net_row, start_column=1, end_row=net_row, end_column=6)
    ws.cell(net_row, 1, "стоимость без НДС").alignment = Alignment(horizontal="right")
    ws.cell(net_row, 7, NET).number_format = '#,##0.00'

    foot = net_row + 2
    ws.merge_cells(start_row=foot, start_column=1, end_row=foot, end_column=7)
    ws.cell(foot, 1, "Условия оплаты: предоплата 100%. Цены указаны с учётом стоимости доставки.")
    ws.merge_cells(start_row=foot + 1, start_column=1, end_row=foot + 1, end_column=7)
    ws.cell(
        foot + 1,
        1,
        "Шкаф 12U и ИБП 2000 ВА рассчитаны на данный комплект (24p HiPoE + 16ch NVR + 2 HDD + PoE-нагрузка 16 камер).",
    )

    widths = [5, 62, 28, 10, 10, 16, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[header_row].height = 22
    for i in range(1, len(ITEMS) + 1):
        ws.row_dimensions[header_row + i].height = 36
    ws.print_title_rows = f"1:{header_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.freeze_panes = f"A{header_row + 1}"

    # --- Сравнение ---
    ws2 = wb.create_sheet("Сравнение с исходным КП")
    ws2["A1"] = "Корректировки относительно исходного КП ООО «Эффорт Юг»"
    ws2["A1"].font = font_title
    ws2["A1"].fill = fill_navy
    ws2.merge_cells("A1:F1")
    ws2.row_dimensions[1].height = 28

    cmp_headers = ["Позиция ТЗ", "Было в КП", "Стало", "Кол-во было", "Кол-во стало", "Почему"]
    for col, h in enumerate(cmp_headers, 1):
        c = ws2.cell(3, col, h)
        c.font = font_white
        c.fill = fill_head
        c.alignment = center
        c.border = thin

    cmp_rows = [
        (
            "Камеры 4 Мп 2.8 мм ColorVu",
            "DS-2CD2043G2-IU — AcuSense, ИК, не ColorVu",
            "DS-2CD2047G2-LU(C)(2.8mm) ColorVu",
            5,
            16,
            "Другая технология изображения; нехватка количества",
        ),
        (
            "Монтажные коробки",
            "не было",
            "DS-1280ZJ-S",
            0,
            16,
            "Коммутация кабеля, герметизация, аккуратный монтаж",
        ),
        (
            "Регистратор 16 каналов",
            "DS-7608NXI-K2/8P(E) — 8 кан. / 8 PoE",
            "DS-7616NXI-K2 — 16 кан., без PoE",
            1,
            1,
            "8 каналов нельзя расширить до 16; PoE вынесен на коммутатор",
        ),
        (
            "HDD на 2 месяца",
            "ST2000VX017 2 ТБ — ≈2–8 суток",
            "2×12 ТБ SkyHawk ≈ 60 суток H.265+",
            1,
            2,
            "Расчёт 16×2,1 Мбит/с × 60 сут ≈ 21,8 ТБ + 10%",
        ),
        (
            "Коммутатор HiPoE 24 порта",
            "нет (PoE на 8-канальном NVR)",
            "DS-3E1326P-EI 24p / 370 Вт",
            0,
            1,
            "16 камер + запас 8 портов; бюджет 370 Вт при нагрузке 120 Вт",
        ),
        (
            "Шкаф",
            "нет",
            "12U 19\" 600×600 настенный",
            0,
            1,
            "Расчёт под коммутатор 1U + NVR + ИБП-башня + PDU",
        ),
        (
            "ИБП",
            "нет",
            "2000 ВА чистая синусоида",
            0,
            1,
            "Нагрузка ~200 Вт (коммутатор+NVR+PoE камер), автономия 15–25 мин",
        ),
    ]
    for i, rowv in enumerate(cmp_rows, 1):
        for col, val in enumerate(rowv, 1):
            c = ws2.cell(3 + i, col, val)
            c.font = font_n
            c.alignment = wrap
            c.border = thin
            if i % 2 == 0:
                c.fill = fill_alt
        ws2.row_dimensions[3 + i].height = 48

    ws2.cell(12, 1, "Итого исходное КП, ₽").font = font_bold
    ws2.cell(12, 2, 62310).number_format = '#,##0'
    ws2.cell(13, 1, "Итого скорректированное КП, ₽").font = font_bold
    ws2.cell(13, 2, TOTAL).number_format = '#,##0'
    ws2.cell(13, 2).fill = fill_total

    for i, w in enumerate([28, 42, 42, 14, 14, 48], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.page_setup.orientation = "landscape"
    ws2.page_setup.fitToPage = True
    ws2.page_setup.fitToWidth = 1
    ws2.page_setup.fitToHeight = 1

    # --- Расчёт HDD ---
    ws3 = wb.create_sheet("Расчёт архива 60 суток")
    ws3["A1"] = "Архив 60 суток: 16 камер 4 Мп ColorVu, H.265+"
    ws3["A1"].font = font_title
    ws3["A1"].fill = fill_navy
    ws3.merge_cells("A1:E1")
    calc_h = ["Сценарий", "Битрейт на камеру", "В сутки, ГБ", "За 60 суток, ТБ", "Диски"]
    for col, h in enumerate(calc_h, 1):
        c = ws3.cell(3, col, h)
        c.font = font_white
        c.fill = fill_head
        c.border = thin
        c.alignment = center
    calc_data = [
        ("Рабочий: H.265+ среднее + аудио (заложен в КП)", "2,1 Мбит/с", 363, 21.8, "2×12 ТБ"),
        ("Высокое качество", "3,0 Мбит/с", 518, 31.1, "2×16 ТБ (макс. NVR)"),
        ("Только детекция ~50%", "эфф. 1,05 Мбит/с", 181, 10.9, "не рекомендуется для аптеки"),
        ("Исходный 1×2 ТБ на 16 камер", "2,1 Мбит/с", 363, 1.8, "≈ 2,5 суток — не подходит"),
    ]
    for i, rowv in enumerate(calc_data, 1):
        for col, val in enumerate(rowv, 1):
            c = ws3.cell(3 + i, col, val)
            c.border = thin
            c.alignment = wrap
            if i == 1:
                c.fill = PatternFill("solid", fgColor="C6EFCE")
            elif i == 4:
                c.fill = PatternFill("solid", fgColor="FFC7CE")
        ws3.row_dimensions[3 + i].height = 28
    note = (
        "Формула: ГБ/сут = Мбит/с × 10,8 × число камер. "
        "16 × 2,1 × 10,8 = 363 ГБ/сут; 363 × 60 / 1024 ≈ 21,8 ТБ. "
        "+10% индексы ≈ 24 ТБ = 2×12 ТБ. DS-7616NXI-K2: 2 SATA, до 16 ТБ на диск."
    )
    ws3.merge_cells("A8:E9")
    ws3["A8"] = note
    ws3["A8"].alignment = wrap
    for i, w in enumerate([48, 22, 16, 18, 32], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    # --- Питание ---
    ws4 = wb.create_sheet("Шкаф и ИБП")
    ws4["A1"] = "Расчёт шкафа и ИБП на комплект аптеки"
    ws4["A1"].font = font_title
    ws4["A1"].fill = fill_navy
    ws4.merge_cells("A1:D1")
    p_h = ["Устройство", "Типично, Вт", "Максимум, Вт", "Примечание"]
    for col, h in enumerate(p_h, 1):
        c = ws4.cell(3, col, h)
        c.font = font_white
        c.fill = fill_head
        c.border = thin
    power = [
        ("16× камера ColorVu PoE 802.3af", 80, 120, "7,5 Вт max, Hi-PoE 60 Вт не нужен"),
        ("DS-3E1326P-EI (свой расход + PoE)", 115, 155, "бюджет порта 370 Вт, запас 68%"),
        ("DS-7616NXI-K2 + 2 HDD", 27, 40, "старт шпинделя — пик"),
        ("Вентиляция / PDU", 10, 15, ""),
        ("ИТОГО на ИБП", 210, 250, "ИБП 2000 ВА / ≥1200 Вт, чистая синусоида"),
    ]
    for i, rowv in enumerate(power, 1):
        for col, val in enumerate(rowv, 1):
            c = ws4.cell(3 + i, col, val)
            c.border = thin
            c.alignment = wrap
            if i == 5:
                c.font = font_bold
                c.fill = fill_total
        ws4.row_dimensions[3 + i].height = 24
    ws4["A10"] = "Шкаф"
    ws4["A10"].font = font_bold
    ws4.merge_cells("A11:D13")
    ws4["A11"] = (
        "12U 19\" 600×600 настенный: 1U коммутатор + 2U полка NVR (глубина корпуса 315 мм) "
        "+ 1U PDU + органайзеры + башенный ИБП 2000 ВА на дно (~15–18 кг). "
        "Оборудование ≈ 28 кг при нагрузке шкафа от 60 кг. "
        "9U 450 мм недостаточен по глубине под NVR+ИБП."
    )
    ws4["A11"].alignment = wrap
    for i, w in enumerate([42, 16, 16, 48], 1):
        ws4.column_dimensions[get_column_letter(i)].width = w

    # --- Источники ---
    ws5 = wb.create_sheet("Источники цен")
    ws5["A1"] = "Ориентиры рынка (август 2026) и цена в КП"
    ws5["A1"].font = font_title
    ws5["A1"].fill = fill_navy
    ws5.merge_cells("A1:D1")
    src_h = ["Позиция", "Ориентир рынка с НДС, ₽", "Цена в КП, ₽", "Источник / логика"]
    for col, h in enumerate(src_h, 1):
        c = ws5.cell(3, col, h)
        c.font = font_white
        c.fill = fill_head
        c.border = thin
        c.alignment = center
    sources = [
        (
            "DS-2CD2047G2-LU(C)(2.8mm)",
            "12 654 … 21 090",
            12990,
            "hikvisionpro.ru 21 090; emart.su акция 12 654. В исходном КП AcuSense была 6 670 при рознице ~10–11 тыс. Та же логика дилерской цены.",
        ),
        (
            "DS-1280ZJ-S",
            "1 328 … 3 090",
            1290,
            "Softline 1 328, secbuy 1 350, DSSL 3 090.",
        ),
        (
            "DS-7616NXI-K2",
            "20 113 … 30 790",
            22490,
            "videoglaz.ru 20 113, us-plast 30 790. Исходный 8ch PoE был 15 570.",
        ),
        (
            "SkyHawk 12 ТБ",
            "27 870 … 40 135",
            28900,
            "ST12000VE001: Планета компьютеров 27 870, XPS-PRO 28 737.",
        ),
        (
            "DS-3E1326P-EI",
            "30 690 … 33 119",
            32990,
            "Нева Электроникс 30 690, DSSL 33 119. 24p PoE 370 Вт (HiPoE-бюджет).",
        ),
        (
            "Шкаф 12U 600×600",
            "11 000 … 20 530",
            16900,
            "Настенные 9–12U ЦМО / МиК / Tinko.",
        ),
        (
            "ИБП 2000 ВА чистый синус",
            "25 000 … 35 000",
            28900,
            "Линейно-интерактив, нагрузка ~200 Вт, 15–25 мин.",
        ),
    ]
    for i, rowv in enumerate(sources, 1):
        for col, val in enumerate(rowv, 1):
            c = ws5.cell(3 + i, col, val)
            c.border = thin
            c.alignment = wrap
            if col == 3:
                c.number_format = '#,##0'
        ws5.row_dimensions[3 + i].height = 40
    ws5.merge_cells("A12:D13")
    ws5["A12"] = (
        "Цены КП — ориентир поставки с НДС 22% в логике исходного КП Эффорт Юг, не оферта розничного магазина. "
        "Перед счётом уточнить наличие ревизий (C)/(D) и текущий артикул SkyHawk 12 ТБ."
    )
    ws5["A12"].alignment = wrap
    for i, w in enumerate([32, 22, 16, 78], 1):
        ws5.column_dimensions[get_column_letter(i)].width = w

    path = ROOT / "КП_Аптека_2026-08-18.xlsx"
    wb.save(path)
    return path


def build_docx() -> Path:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.6)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)

    p = doc.add_paragraph()
    r = p.add_run("ООО «Эффорт Юг»")
    set_run_font(r, "Calibri", 16, True, NAVY_RGB)
    p.paragraph_format.space_after = Pt(0)

    for line in [
        "ИНН 6165231466, ОГРН 1226100001148",
        "344064, Ростовская обл., г. Ростов-на-Дону,",
        "ул. Вавилова, д. 71 В стр. 2, этаж 2 помещ. 2 ж",
    ]:
        p = doc.add_paragraph()
        set_run_font(p.add_run(line), "Calibri", 10, False, RGBColor(0x5B, 0x5B, 0x5B))
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t.paragraph_format.space_before = Pt(12)
    set_run_font(t.add_run("КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ"), "Calibri", 18, True, NAVY_RGB)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(sub.add_run(f"{KP_NO} от {TODAY.strftime('%d.%m.%Y')}"), "Calibri", 12, True)

    info = doc.add_paragraph()
    set_run_font(
        info.add_run(
            "Объект: аптека. Система видеонаблюдения 16 камер 4 Мп ColorVu.\n"
            "Корректировка исходного комплекта (8 каналов, 5 камер AcuSense, HDD 2 ТБ)."
        ),
        "Calibri",
        11,
    )

    table = doc.add_table(rows=1, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    hdr = ["№", "Наименование", "Ед.", "Кол-во", "Цена, руб.", "Сумма, руб."]
    for i, h in enumerate(hdr):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_font(p.add_run(h), "Calibri", 10, True, RGBColor(255, 255, 255))
        shade_cell(cell, "1F4E79")
        set_cell_border(cell)

    for i, item in enumerate(ITEMS, 1):
        row = table.add_row().cells
        vals = [
            str(i),
            item["name"],
            item["unit"],
            str(item["qty"]),
            money_ru(item["price"]),
            money_ru(item["sum"]),
        ]
        aligns = [
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ]
        for col, (val, al) in enumerate(zip(vals, aligns)):
            row[col].text = ""
            p = row[col].paragraphs[0]
            p.alignment = al
            set_run_font(p.add_run(val), "Calibri", 9, col == 0)
            if i % 2 == 0:
                shade_cell(row[col], "D6EAF8")
            set_cell_border(row[col])

    tot = table.add_row().cells
    tot[0].merge(tot[4])
    tot[0].text = ""
    p = tot[0].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_run_font(p.add_run("Итого общая стоимость оборудования с НДС 22%"), "Calibri", 10, True)
    tot[5].text = ""
    p = tot[5].paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_run_font(p.add_run(money_ru(TOTAL)), "Calibri", 11, True)
    for c in tot:
        shade_cell(c, "FFF2CC")
        set_cell_border(c)

    for line in [
        f"в том числе НДС 22% — {money_ru(VAT)} руб.",
        f"стоимость без НДС — {money_ru(NET)} руб.",
        "Условия оплаты: предоплата 100%.",
        "Цены указаны с учётом стоимости доставки.",
        "Шкаф и ИБП рассчитаны на данный комплект (коммутатор 24p HiPoE, регистратор 16ch, 2 HDD, PoE-нагрузка 16 камер).",
    ]:
        p = doc.add_paragraph()
        set_run_font(p.add_run(line), "Calibri", 11, "Итого" in line or "НДС" in line)
        p.paragraph_format.space_after = Pt(2)

    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(12)
    set_run_font(h.add_run("Состав относительно исходного КП"), "Calibri", 13, True, NAVY_RGB)

    bullets = [
        "Камеры: 5× AcuSense ИК → 16× ColorVu 4 Мп 2.8 мм + 16 монтажных коробок.",
        "Регистратор: 8 каналов с PoE → 16 каналов без PoE (питание камер с коммутатора).",
        "Архив: 1×2 ТБ (~двое–восемь суток) → 2×12 ТБ (~60 суток при H.265+).",
        "Сеть: добавлен 24-портовый PoE-коммутатор 370 Вт (HiPoE-бюджет).",
        "Инфраструктура: шкаф 12U 600×600 и ИБП 2000 ВА с чистой синусоидой.",
    ]
    for b in bullets:
        p = doc.add_paragraph(style="List Bullet")
        p.clear()
        set_run_font(p.add_run(b), "Calibri", 11)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    set_run_font(p.add_run("Заказчик / объект: аптека."), "Calibri", 11)

    path = ROOT / "КП_Аптека_2026-08-18.docx"
    doc.save(path)
    return path


def build_pdf() -> Path:
    pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
    pdfmetrics.registerFont(TTFont("DejaVuBold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))

    path = ROOT / "КП_Аптека_2026-08-18.pdf"
    doc = SimpleDocTemplate(
        str(path),
        pagesize=A4,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=12 * mm,
        bottomMargin=14 * mm,
        title=KP_NO,
        author="ООО «Эффорт Юг»",
    )
    styles = getSampleStyleSheet()
    s_company = ParagraphStyle(
        "co", parent=styles["Normal"], fontName="DejaVuBold", fontSize=13, textColor=colors.HexColor("#1F4E79"), leading=16
    )
    s_small = ParagraphStyle(
        "sm", parent=styles["Normal"], fontName="DejaVu", fontSize=8, textColor=colors.HexColor("#555555"), leading=11
    )
    s_title = ParagraphStyle(
        "ti", parent=styles["Normal"], fontName="DejaVuBold", fontSize=14, alignment=TA_CENTER, textColor=colors.HexColor("#1F4E79"), spaceBefore=8, spaceAfter=4
    )
    s_center = ParagraphStyle(
        "ce", parent=styles["Normal"], fontName="DejaVuBold", fontSize=10, alignment=TA_CENTER, spaceAfter=8
    )
    s_body = ParagraphStyle(
        "bd", parent=styles["Normal"], fontName="DejaVu", fontSize=9, leading=12, alignment=TA_JUSTIFY, spaceAfter=6
    )
    s_h = ParagraphStyle(
        "hh", parent=styles["Normal"], fontName="DejaVuBold", fontSize=11, textColor=colors.HexColor("#1F4E79"), spaceBefore=10, spaceAfter=4
    )
    s_cell = ParagraphStyle("cl", parent=styles["Normal"], fontName="DejaVu", fontSize=7.5, leading=10)
    s_cell_c = ParagraphStyle("clc", parent=s_cell, alignment=TA_CENTER)
    s_cell_r = ParagraphStyle("clr", parent=s_cell, alignment=TA_RIGHT)
    s_cell_b = ParagraphStyle("clb", parent=s_cell, fontName="DejaVuBold", alignment=TA_RIGHT)
    s_th = ParagraphStyle(
        "th", parent=styles["Normal"], fontName="DejaVuBold", fontSize=7.5, textColor=colors.white, alignment=TA_CENTER, leading=10
    )
    s_foot = ParagraphStyle("ft", parent=styles["Normal"], fontName="DejaVu", fontSize=8, leading=11, spaceBefore=2)

    story = []
    story.append(Paragraph("ООО «Эффорт Юг»", s_company))
    story.append(Paragraph("ИНН 6165231466, ОГРН 1226100001148", s_small))
    story.append(Paragraph("344064, Ростовская обл., г. Ростов-на-Дону, ул. Вавилова, д. 71 В стр. 2, этаж 2 помещ. 2 ж", s_small))
    story.append(Paragraph("КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ", s_title))
    story.append(Paragraph(f"{KP_NO} от {TODAY.strftime('%d.%m.%Y')}", s_center))
    story.append(
        Paragraph(
            "Объект: аптека. Система видеонаблюдения на 16 камер 4 Мп ColorVu. "
            "Документ заменяет исходный комплект (регистратор 8 каналов, 5 камер AcuSense, HDD 2 ТБ).",
            s_body,
        )
    )

    header = [
        Paragraph("№", s_th),
        Paragraph("Наименование", s_th),
        Paragraph("Ед.", s_th),
        Paragraph("Кол-во", s_th),
        Paragraph("Цена, руб.", s_th),
        Paragraph("Сумма, руб.", s_th),
    ]
    data = [header]
    for i, item in enumerate(ITEMS, 1):
        data.append(
            [
                Paragraph(str(i), s_cell_c),
                Paragraph(item["name"], s_cell),
                Paragraph(item["unit"], s_cell_c),
                Paragraph(str(item["qty"]), s_cell_c),
                Paragraph(money_ru(item["price"]), s_cell_r),
                Paragraph(money_ru(item["sum"]), s_cell_r),
            ]
        )
    data.append(
        [
            Paragraph("", s_cell),
            Paragraph("Итого общая стоимость оборудования с НДС 22%", s_cell_b),
            Paragraph("", s_cell),
            Paragraph("", s_cell),
            Paragraph("", s_cell),
            Paragraph(money_ru(TOTAL), s_cell_b),
        ]
    )

    tbl = Table(data, colWidths=[10 * mm, 96 * mm, 16 * mm, 16 * mm, 26 * mm, 26 * mm])
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#FFF2CC")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("SPAN", (1, -1), (4, -1)),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#8FAADC")),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    for i in range(1, len(ITEMS) + 1):
        if i % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#EAF3FB")))
    tbl.setStyle(TableStyle(style_cmds))
    story.append(tbl)
    story.append(Spacer(1, 6))
    story.append(Paragraph(f"в том числе НДС 22% — {money_ru(VAT)} руб. &nbsp;&nbsp; стоимость без НДС — {money_ru(NET)} руб.", s_foot))
    story.append(Paragraph("Условия оплаты: предоплата 100%. Цены указаны с учётом стоимости доставки.", s_foot))
    story.append(
        Paragraph(
            "Шкаф 12U и ИБП 2000 ВА рассчитаны на данный комплект: коммутатор 24p HiPoE, регистратор 16ch, два диска, PoE-нагрузка 16 камер.",
            s_foot,
        )
    )

    story.append(Paragraph("Что изменено относительно исходного КП", s_h))
    story.append(
        Paragraph(
            "• Камеры: 5 шт. DS-2CD2043G2-IU (AcuSense, ИК, не ColorVu) → 16 шт. DS-2CD2047G2-LU(C) 4 Мп ColorVu 2.8 мм + 16 коробок DS-1280ZJ-S.<br/>"
            "• Регистратор: DS-7608NXI-K2/8P(E) 8 каналов / 8 PoE → DS-7616NXI-K2 16 каналов без PoE.<br/>"
            "• Архив: 1×2 ТБ (примерно 2–8 суток) → 2×12 ТБ, около 60 суток при H.265+ 2,1 Мбит/с.<br/>"
            "• Сеть: добавлен Hikvision DS-3E1326P-EI, 24 порта PoE, бюджет 370 Вт. Камерам достаточно 802.3af (7,5 Вт); Hi-PoE 60 Вт для PTZ не требуется, нужен высокий суммарный бюджет на 24 порта.<br/>"
            "• Шкаф 12U 600×600 и ИБП 2000 ВА с чистой синусоидой — под фактическую нагрузку ~200 Вт (15–25 мин автономии).",
            s_body,
        )
    )

    story.append(Paragraph("Расчёт архива 60 суток", s_h))
    arch_header = [
        Paragraph(x, s_th)
        for x in ["Сценарий", "Битрейт", "За 60 суток", "Диски"]
    ]
    arch = [arch_header]
    for row in [
        ("Рабочий (заложен в КП): H.265+ среднее + аудио", "2,1 Мбит/с × 16", "21,8 ТБ (+10% → 24 ТБ)", "2×12 ТБ"),
        ("Высокое качество", "3,0 Мбит/с × 16", "31,1 ТБ", "2×16 ТБ, предел NVR"),
        ("Исходный ST2000VX017 на 16 камер", "2,1 Мбит/с × 16", "≈ 2,5 суток", "не подходит"),
    ]:
        arch.append([Paragraph(x, s_cell) for x in row])
    t2 = Table(arch, colWidths=[70 * mm, 40 * mm, 45 * mm, 35 * mm])
    t2.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2E75B6")),
                ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#C6EFCE")),
                ("BACKGROUND", (0, 3), (-1, 3), colors.HexColor("#FFC7CE")),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#8FAADC")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    story.append(t2)
    story.append(
        Paragraph(
            "Формула: 16 камер × 2,1 Мбит/с × 10,8 ГБ/сут × 60 сут = 21,8 ТБ. Регистратор DS-7616NXI-K2: 2 SATA, до 16 ТБ на диск.",
            s_foot,
        )
    )

    story.append(Paragraph("Шкаф и ИБП — расчёт на комплект", s_h))
    story.append(
        Paragraph(
            "Нагрузка типичная ≈ 190–210 Вт, пик ≈ 250 Вт (старт дисков). Камеры 16×7,5 Вт = 120 Вт PoE при бюджете коммутатора 370 Вт. "
            "ИБП 2000 ВА / ≥1200 Вт, чистая синусоида — запас на PFC-блок коммутатора. "
            "Шкаф 12U 19\" 600×600: коммутатор 1U, полка под NVR (корпус 315 мм), PDU, башенный ИБП ~16 кг на дно. Масса оборудования ≈ 28 кг.",
            s_body,
        )
    )
    story.append(Paragraph("Заказчик / объект: аптека.", s_foot))

    def footer(canvas, doc_):
        canvas.saveState()
        canvas.setFont("DejaVu", 8)
        canvas.setFillColor(colors.HexColor("#666666"))
        canvas.drawString(14 * mm, 8 * mm, "ООО «Эффорт Юг»  ·  не является публичной офертой, цены ориентировочные")
        canvas.drawRightString(A4[0] - 14 * mm, 8 * mm, f"стр. {doc_.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return path


def main():
    xlsx = build_xlsx()
    docx = build_docx()
    pdf = build_pdf()
    print(f"TOTAL={TOTAL} VAT={VAT} NET={NET}")
    print(xlsx)
    print(docx)
    print(pdf)


if __name__ == "__main__":
    main()
