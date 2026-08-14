#!/usr/bin/env python3
"""Ориентировочное КП ред. 2: 2× NeuroStation + 1× 1143 + 49× 4MBIR-28-TMLW + 1× PTZ."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

VAT_RATE = 0.22
DATE = "14.08.2026"
VALID_UNTIL = "28.08.2026"
KP_NO = "КП-2026-08-14-001"

NAVY = "1F4E79"
NAVY2 = "2E75B6"
LIGHT = "D6E4F0"
GOLD = "FFF2CC"
GREEN = "E2EFDA"
ORANGE = "FCE4D6"
WHITE = "FFFFFF"
GRAY = "F2F2F2"
RED_NOTE = "C00000"

THIN = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)


def vat_from_gross(gross: float) -> float:
    return round(gross * VAT_RATE / (1 + VAT_RATE), 2)


def net_from_gross(gross: float) -> float:
    return round(gross - vat_from_gross(gross), 2)


ITEMS_REQUEST = [
    {
        "no": 1,
        "name": (
            "IP-видеорегистратор TRASSIR NeuroStation Astra 9800R/128-S, "
            "128 каналов, Astra Linux SE «Смоленск», 8×HDD 3.5\" (диски не входят), "
            "2 лицензии Neuro Detector в комплекте. СТ-1 / реестры ПП №719 и №878"
        ),
        "sku": "NeuroStation 9800R/128-S",
        "unit": "шт.",
        "qty": 2,
        "price": 477740.00,
        "note": "Открытая розница trasrussia.ru; у DSSL — проектная цена",
    },
    {
        "no": 2,
        "name": (
            "Купольная IP-камера 4 Мп НИЦ модель 1143, мотор 2.7–13.5 мм, ИК до 35 м, "
            "IP66/IK10, ПО Nexus. Производство РФ, СТ-1 на партию"
        ),
        "sku": "4MP-DOM-2.7-13.5M Nexus",
        "unit": "шт.",
        "qty": 1,
        "price": 23800.00,
        "note": "Спецификация НИЦ PDF; ориентир цены 4MP-DOM-2.7-13.5M; СТ-1 запросить на партию",
    },
    {
        "no": 3,
        "name": (
            "Уличная цилиндрическая IP-камера 4 Мп Mini Bullet, объектив 2.8 мм, "
            "ИК + белый свет, микрофон. Производитель РУВЕР, СТ-1"
        ),
        "sku": "4MBIR-28-TMLW",
        "unit": "шт.",
        "qty": 49,
        "price": 18500.00,
        "note": "https://www.ipplus.ru/catalog/ip-kamery/ip-kamery-4mp/4mbir-28-tmlw ; st_1=true",
    },
    {
        "no": 4,
        "name": (
            "PTZ уличная 2 Мп, зум ×20 (4.7–94 мм), IP66. Самая недорогая PTZ РУВЕР/Айпи Плюс "
            "с СТ-1 в серии 2 Мп (×10 без СТ-1 не предлагаются)"
        ),
        "sku": "2BPBDD-4794-20",
        "unit": "шт.",
        "qty": 1,
        "price": 92750.00,
        "note": "https://www.ipplus.ru/catalog/ptz-kamery/ptz-kamery-2mp/2bpbdd-4794-20 ; реестр Минпромторга",
    },
]

ITEM_ANYIP = {
    "no": 5,
    "name": "Лицензия TRASSIR AnyIP (TRASSIR OS) — подключение 1 IP-камеры любого производителя",
    "sku": "DSSL 57885",
    "unit": "шт.",
    "qty": 51,
    "price": 5090.00,
    "note": "Цена DSSL с 01.04.2026; без лицензии камера к NVR не подключается",
}

ITEM_HDD = {
    "no": 6,
    "name": "HDD 10 ТБ Seagate SkyHawk AI, 3.5\" SATA, рекомендован DSSL для NeuroStation 8-disk",
    "sku": "ST10000VE001",
    "unit": "шт.",
    "qty": 8,
    "price": 55200.00,
    "note": "Ориентир ANDPRO; ~30 суток архива при 51 камере / 3 Мбит/с",
}

ITEM_ND = {
    "no": 7,
    "name": "TRASSIR Neuro Detector — доп. каналы (в 2 NVR уже 4 лицензии)",
    "sku": "DSSL 26363",
    "unit": "шт.",
    "qty": 47,
    "price": 6600.00,
    "note": "Опция: аналитика на все 51 камеру",
}


def enrich(item: dict) -> dict:
    out = dict(item)
    out["sum"] = round(out["price"] * out["qty"], 2)
    out["vat"] = vat_from_gross(out["sum"])
    out["net"] = net_from_gross(out["sum"])
    return out


def money(n: float) -> float:
    return round(n, 2)


def header_cell(ws, row, col, value, fill=NAVY, size=10, wrap=True):
    cell = ws.cell(row, col, value)
    cell.font = Font(bold=True, color=WHITE, size=size, name="Calibri")
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    cell.border = THIN
    return cell


def apply_range_border(ws, r1, r2, c1, c2):
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for cell in row:
            cell.border = THIN
            if cell.font is None or cell.font.name is None:
                cell.font = Font(name="Calibri", size=10)


def write_item_row(ws, row, item, fill=None):
    values = [
        item["no"],
        item["name"],
        item["sku"],
        item["unit"],
        item["qty"],
        item["price"],
        item["sum"],
        item["vat"],
        item["note"],
    ]
    for col, val in enumerate(values, 1):
        cell = ws.cell(row, col, val)
        cell.font = Font(name="Calibri", size=10)
        cell.alignment = Alignment(
            vertical="center",
            wrap_text=True,
            horizontal="center" if col in (1, 4, 5) else "left",
        )
        if col in (6, 7, 8):
            cell.number_format = '#,##0.00'
            cell.alignment = Alignment(vertical="center", horizontal="right")
        if fill:
            cell.fill = PatternFill("solid", fgColor=fill)
        cell.border = THIN
    ws.row_dimensions[row].height = 48


def totals_block(ws, start_row, label, gross, fill):
    net = net_from_gross(gross)
    vat = vat_from_gross(gross)
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=6)
    cell = ws.cell(start_row, 1, label)
    cell.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.alignment = Alignment(horizontal="right", vertical="center")
    for col in range(1, 7):
        ws.cell(start_row, col).fill = PatternFill("solid", fgColor=fill)
        ws.cell(start_row, col).border = THIN
        ws.cell(start_row, col).font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    for col, val in ((7, gross), (8, vat)):
        c = ws.cell(start_row, col, val)
        c.number_format = '#,##0.00'
        c.font = Font(name="Calibri", size=11, bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.border = THIN
    ws.cell(start_row, 9).fill = PatternFill("solid", fgColor=fill)
    ws.cell(start_row, 9).border = THIN
    ws.row_dimensions[start_row].height = 22
    return net, vat


def build_kp_sheet(ws):
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.print_title_rows = "1:8"

    widths = {1: 6, 2: 52, 3: 26, 4: 8, 5: 10, 6: 16, 7: 16, 8: 14, 9: 42}
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.merge_cells("A1:I1")
    t = ws["A1"]
    t.value = "КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ"
    t.font = Font(name="Calibri", size=18, bold=True, color=WHITE)
    t.fill = PatternFill("solid", fgColor=NAVY)
    t.alignment = Alignment(horizontal="center", vertical="center")
    for col in range(1, 10):
        ws.cell(1, col).fill = PatternFill("solid", fgColor=NAVY)
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:I2")
    ws["A2"] = (
        f"№ {KP_NO} ред. 2 от {DATE}  ·  2× NVR TRASSIR + 1× НИЦ 1143 + 49× 4MBIR-28-TMLW + 1× PTZ  "
        f"·  все камеры со СТ-1  ·  цены с НДС 22%"
    )
    ws["A2"].font = Font(name="Calibri", size=10, italic=True, color=WHITE)
    ws["A2"].fill = PatternFill("solid", fgColor=NAVY2)
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    for col in range(1, 10):
        ws.cell(2, col).fill = PatternFill("solid", fgColor=NAVY2)
    ws.row_dimensions[2].height = 18

    meta = [
        ("Заказчик", "______________________________"),
        ("ИНН / КПП заказчика", "______________________________"),
        ("Поставщик", "к заполнению (DSSL / TRASSIR, НИЦ «Технологии», ООО «Айпи Плюс» / РУВЕР)"),
        ("Срок действия цен", VALID_UNTIL),
        ("Срок поставки", "NVR: 5–15 раб. дней при наличии; камеры НИЦ и РУВЕР — под заказ"),
        ("Условия оплаты", "по договору (типовой ориентир: 100% предоплата или 70/30)"),
    ]
    r = 4
    for label, value in meta:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        a = ws.cell(r, 1, label)
        a.font = Font(name="Calibri", size=10, bold=True)
        a.fill = PatternFill("solid", fgColor=LIGHT)
        ws.cell(r, 2).fill = PatternFill("solid", fgColor=LIGHT)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=9)
        b = ws.cell(r, 3, value)
        b.font = Font(name="Calibri", size=10)
        r += 1

    ws.merge_cells("A11:I11")
    ws["A11"] = "1. Спецификация по запросу заказчика"
    ws["A11"].font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    ws["A11"].fill = PatternFill("solid", fgColor=NAVY)
    for col in range(1, 10):
        ws.cell(11, col).fill = PatternFill("solid", fgColor=NAVY)
    ws.row_dimensions[11].height = 20

    headers = [
        "№",
        "Наименование",
        "Артикул",
        "Ед.",
        "Кол-во",
        "Цена с НДС, ₽",
        "Сумма с НДС, ₽",
        "в т.ч. НДС 22%, ₽",
        "Примечание",
    ]
    for col, h in enumerate(headers, 1):
        header_cell(ws, 12, col, h)
    ws.row_dimensions[12].height = 28

    request_items = [enrich(x) for x in ITEMS_REQUEST]
    fills = [GRAY, WHITE, GRAY, WHITE]
    for i, item in enumerate(request_items):
        write_item_row(ws, 13 + i, item, fills[i])
    request_sum = money(sum(x["sum"] for x in request_items))
    totals_block(ws, 17, "Итого по запросу (позиции 1–4), камеры 51 шт. со СТ-1", request_sum, NAVY)

    ws.merge_cells("A19:I19")
    ws["A19"] = (
        "2. Обязательно к поставке для работоспособности "
        "(в регистраторе нет лицензий на IP-камеры; камеры НИЦ/РУВЕР не native TRASSIR)"
    )
    ws["A19"].font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    ws["A19"].fill = PatternFill("solid", fgColor=NAVY)
    for col in range(1, 10):
        ws.cell(19, col).fill = PatternFill("solid", fgColor=NAVY)
    ws.row_dimensions[19].height = 20

    for col, h in enumerate(headers, 1):
        header_cell(ws, 20, col, h)
    anyip = enrich(ITEM_ANYIP)
    write_item_row(ws, 21, anyip, GOLD)
    work_sum = money(request_sum + anyip["sum"])
    totals_block(ws, 22, "Итого работоспособный минимум (позиции 1–5)", work_sum, "375623")

    ws.merge_cells("A24:I24")
    ws["A24"] = "3. Рекомендуемые опции (в исходном запросе отсутствуют)"
    ws["A24"].font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    ws["A24"].fill = PatternFill("solid", fgColor=NAVY)
    for col in range(1, 10):
        ws.cell(24, col).fill = PatternFill("solid", fgColor=NAVY)

    for col, h in enumerate(headers, 1):
        header_cell(ws, 25, col, h)
    hdd = enrich(ITEM_HDD)
    nd = enrich(ITEM_ND)
    write_item_row(ws, 26, hdd, GREEN)
    write_item_row(ws, 27, nd, ORANGE)

    archive_sum = money(work_sum + hdd["sum"])
    full_sum = money(archive_sum + nd["sum"])
    totals_block(ws, 28, "Итого с архивом ~30 суток (позиции 1–6)", archive_sum, "548235")
    totals_block(ws, 29, "Итого полный комплект с аналитикой (позиции 1–7)", full_sum, NAVY)

    ws.merge_cells("A31:I31")
    ws["A31"] = "4. Сводка комплектов"
    ws["A31"].font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    ws["A31"].fill = PatternFill("solid", fgColor=NAVY)
    for col in range(1, 10):
        ws.cell(31, col).fill = PatternFill("solid", fgColor=NAVY)

    summary_headers = ["Комплект", "Состав", "Сумма с НДС, ₽", "в т.ч. НДС 22%, ₽", "Без НДС, ₽"]
    for col, h in enumerate(summary_headers, 1):
        header_cell(ws, 32, col, h)
    ws.merge_cells("E32:I32")
    for col in range(5, 10):
        ws.cell(32, col).fill = PatternFill("solid", fgColor=NAVY)
        ws.cell(32, col).border = THIN

    packs = [
        ("А. Только запрошенные позиции", "2 NVR + 1 купол + 49 цилиндр + 1 PTZ", request_sum, LIGHT),
        ("Б. Работоспособный минимум", "А + 51× AnyIP", work_sum, GOLD),
        ("В. С архивом 30 суток", "Б + 8× HDD 10 ТБ", archive_sum, GREEN),
        ("Г. Полный комплект", "В + Neuro Detector на все камеры", full_sum, ORANGE),
    ]
    for i, (name, composition, gross, fill) in enumerate(packs):
        row = 33 + i
        ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=9)
        ws.cell(row, 1, name).font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row, 2, composition).font = Font(name="Calibri", size=10)
        for col, val in ((3, gross), (4, vat_from_gross(gross)), (5, net_from_gross(gross))):
            c = ws.cell(row, col, val)
            c.number_format = '#,##0.00'
            c.font = Font(name="Calibri", size=10, bold=(i == 1))
            c.alignment = Alignment(horizontal="right", vertical="center")
        for col in range(1, 10):
            ws.cell(row, col).fill = PatternFill("solid", fgColor=fill)
            ws.cell(row, col).border = THIN
            ws.cell(row, col).alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[row].height = 20

    notes = [
        "5. Условия и оговорки",
        "• Документ — ориентировочный расчёт по открытым ценам на 14.08.2026, ред. 2. Не оферта. Фиксация — в КП DSSL, НИЦ и ООО «Айпи Плюс».",
        "• Все камеры поставляются со СТ-1 (и/или номером реестровой записи РЭП). 4MBIR-28-TMLW и 2BPBDD-4794-20: флаг СТ-1 в каталоге IP Plus.",
        "• PTZ: взята самая недорогая модель РУВЕР 2 Мп с СТ-1 — 2BPBDD-4794-20 (зум ×20). Модели ×10 в каталоге без СТ-1 не включены.",
        "• Регистратор без HDD и без лицензий AnyIP. В каждом NVR — 2 лицензии Neuro Detector.",
        "• Оценка архива: 51 камера, H.265, 3 Мбит/с ≈ 1,65 ТБ/сутки; 8×10 ТБ ≈ 30 суток на одном NVR.",
        "• Не включено: PoE-коммутаторы, СКС, шкаф 19\", ИБП, мониторы, монтаж, ПНР, доставка.",
        "• Контакты: DSSL 8 (800) 100-91-12; НИЦ 8 (800) 555-47-65; Айпи Плюс 8 804 333-73-02, info@ipplus.ru.",
    ]
    ws.merge_cells("A38:I38")
    ws["A38"] = notes[0]
    ws["A38"].font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    ws["A38"].fill = PatternFill("solid", fgColor=NAVY)
    for col in range(1, 10):
        ws.cell(38, col).fill = PatternFill("solid", fgColor=NAVY)

    for i, text in enumerate(notes[1:], 39):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=9)
        cell = ws.cell(i, 1, text)
        cell.font = Font(name="Calibri", size=9, color=RED_NOTE if i in (40, 41) else "000000")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[i].height = 18

    ws.merge_cells("A48:D48")
    ws["A48"] = "Поставщик / исполнитель"
    ws["A48"].font = Font(name="Calibri", size=10, bold=True)
    ws.merge_cells("F48:I48")
    ws["F48"] = "Заказчик"
    ws["F48"].font = Font(name="Calibri", size=10, bold=True)
    ws.merge_cells("A50:D50")
    ws["A50"] = "________________ / ________________"
    ws.merge_cells("F50:I50")
    ws["F50"] = "________________ / ________________"
    ws.merge_cells("A51:D51")
    ws["A51"] = "подпись, ФИО, дата"
    ws["A51"].font = Font(name="Calibri", size=8, italic=True, color="808080")
    ws.merge_cells("F51:I51")
    ws["F51"] = "подпись, ФИО, дата"
    ws["F51"].font = Font(name="Calibri", size=8, italic=True, color="808080")

    ws.auto_filter.ref = "A12:I16"
    ws.freeze_panes = "A13"
    ws.print_area = "A1:I51"
    ws.oddHeader.left.text = f"КП {KP_NO}"
    ws.oddFooter.center.text = "Страница &P из &N · ориентировочный расчёт · цены с НДС 22%"

    return {
        "request": request_sum,
        "work": work_sum,
        "archive": archive_sum,
        "full": full_sum,
        "request_vat": vat_from_gross(request_sum),
        "request_net": net_from_gross(request_sum),
        "work_vat": vat_from_gross(work_sum),
        "anyip": anyip["sum"],
        "hdd": hdd["sum"],
        "nd": nd["sum"],
    }


def build_sources_sheet(ws):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 70

    ws.merge_cells("A1:D1")
    ws["A1"] = "Источники цен (снимок 14.08.2026)"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color=WHITE)
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    for col in range(1, 5):
        ws.cell(1, col).fill = PatternFill("solid", fgColor=NAVY)
    ws.row_dimensions[1].height = 24

    headers = ["Позиция", "Артикул", "Цена в КП, ₽", "Источник"]
    for col, h in enumerate(headers, 1):
        header_cell(ws, 3, col, h)
    rows = [
        (
            "NeuroStation Astra 9800R/128-S",
            "NeuroStation 9800R/128-S",
            477740,
            "https://trasrussia.ru/ip-videoregistrator-trassir-neurostation-astra-9800r-128-s",
        ),
        (
            "DSSL карточка NVR (проектная)",
            "80222",
            "проектная",
            "https://www.dssl.ru/products/trassir-neurostation-98128r-s-ip-videoregistrator/",
        ),
        (
            "Купол НИЦ 1143 4MP-DOM-2.7-13.5M Nexus",
            "4MP-DOM-2.7-13.5M",
            23800,
            "https://www.nic-tech.ru/upload/iblock/094/ — спецификация модель 1143",
        ),
        (
            "Цилиндр РУВЕР 4 Мп Mini Bullet СТ-1",
            "4MBIR-28-TMLW",
            18500,
            "https://www.ipplus.ru/catalog/ip-kamery/ip-kamery-4mp/4mbir-28-tmlw (цена на сайте не опубликована)",
        ),
        (
            "PTZ РУВЕР 2 Мп ×20 СТ-1 (самая недорогая с СТ-1)",
            "2BPBDD-4794-20",
            92750,
            "https://www.ipplus.ru/catalog/ptz-kamery/ptz-kamery-2mp/2bpbdd-4794-20 ; ориентир класса PTZ20-20x",
        ),
        (
            "TRASSIR AnyIP",
            "57885",
            5090,
            "https://www.dssl.ru/products/trassir-anyip/",
        ),
        (
            "TRASSIR Neuro Detector",
            "26363",
            6600,
            "https://www.dssl.ru/products/neuro-detector-nejrosetevoj-detector/",
        ),
        (
            "HDD SkyHawk AI 10 ТБ",
            "ST10000VE001",
            55200,
            "https://andpro.ru — 55 202 ₽, в КП округлено 55 200 ₽",
        ),
        (
            "Паспорт NVR",
            "Astra 9800R/128-S",
            "—",
            "https://trassir.ru/products/videonablyudenie/videoregistratory/ip_videoregistratory_nvr/trassir_neurostation_98128r_s/",
        ),
    ]
    for i, (name, sku, price, src) in enumerate(rows, 4):
        ws.cell(i, 1, name).font = Font(name="Calibri", size=10)
        ws.cell(i, 2, sku).font = Font(name="Calibri", size=10)
        cell = ws.cell(i, 3, price)
        cell.font = Font(name="Calibri", size=10)
        if isinstance(price, (int, float)):
            cell.number_format = '#,##0.00'
        ws.cell(i, 4, src).font = Font(name="Calibri", size=9, color="0563C1")
        ws.cell(i, 4).hyperlink = src.split(" — ")[0] if src.startswith("http") else None
        fill = GRAY if i % 2 == 0 else WHITE
        for col in range(1, 5):
            ws.cell(i, col).fill = PatternFill("solid", fgColor=fill)
            ws.cell(i, col).border = THIN
            ws.cell(i, col).alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[i].height = 22


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "КП"
    totals = build_kp_sheet(ws)
    ws2 = wb.create_sheet("Источники цен")
    build_sources_sheet(ws2)
    out = (
        "/workspace/kp/2026-08-14-neurostation-4bhnn0/"
        "КП_поставка_NeuroStation_Astra_9800R_4BHNN0.xlsx"
    )
    wb.save(out)
    print(out)
    for k, v in totals.items():
        print(f"{k}: {v}")


if __name__ == "__main__":
    main()
