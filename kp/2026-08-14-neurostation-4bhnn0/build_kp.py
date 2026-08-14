#!/usr/bin/env python3
"""Ориентировочное КП ред. 4: две заявки ЭТП ГПБ 44-ФЗ (NVR + 4BHNN0)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

VAT_RATE = 0.22
DATE = "14.08.2026"
VALID_UNTIL = "18.08.2026, 00:00 МСК"
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


ITEM_NVR = {
    "no": 1,
    "name": (
        "IP-видеорегистратор TRASSIR NeuroStation Astra 9800R/128-S, "
        "128 каналов, Astra Linux SE «Смоленск», 8×HDD 3.5\" (диски не входят), "
        "2 лицензии Neuro Detector. СТ-1 / реестры ПП №719 и №878"
    ),
    "sku": "NeuroStation 9800R/128-S",
    "unit": "шт.",
    "qty": 2,
    "price": 477740.00,
    "note": "Заявка 254364 / № ЭТП 998818; розница trasrussia.ru; у DSSL проектная",
}

ITEM_CAM = {
    "no": 1,
    "name": (
        "Камера видеонаблюдения цифровая. Предмет извещения — артикул 4BHNN0-0-0-0, "
        "51 шт. одной позиции (эквивалент — только если допущен документацией)"
    ),
    "sku": "4BHNN0-0-0-0",
    "unit": "шт.",
    "qty": 51,
    "price": 22000.00,
    "note": "Заявка 254397 / № ЭТП 998881; артикул в открытых каталогах не найден; ориентир 4 Мп",
}

ITEM_EQUIV = {
    "no": "Э1",
    "name": (
        "Эквивалент (не подавать, если документация не допускает замену): "
        "цилиндр TRASSIR 5 Мп, 2.8 мм, ИК 35 м, микрофон, СТ-1"
    ),
    "sku": "TR-D2151IR3 v2 (R) 2.8",
    "unit": "шт.",
    "qty": 51,
    "price": 16800.00,
    "note": "Только при допуске эквивалента; 51 одинаковых, не 1+49+1",
}

ITEM_ANYIP = {
    "no": 5,
    "name": (
        "Лицензия TRASSIR AnyIP (Astra Linux) — подключение 1 IP-камеры "
        "к NeuroStation Astra / UltraStation Astra"
    ),
    "sku": "DSSL 95887",
    "unit": "шт.",
    "qty": 51,
    "price": 5090.00,
    "note": "Для Astra Linux; цена ориентир AnyIP OS 57885. Подарок на камере — уточнить у DSSL",
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


def section_bar(ws, row, text):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
    ws.cell(row, 1, text).font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    for col in range(1, 10):
        ws.cell(row, col).fill = PatternFill("solid", fgColor=NAVY)
        ws.cell(row, col).border = THIN
        ws.cell(row, col).font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    ws.row_dimensions[row].height = 20


def build_kp_sheet(ws):
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.print_title_rows = "1:3"

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
        f"№ {KP_NO} ред. 4 от {DATE}  ·  44-ФЗ, ЭТП ГПБ «Торговый портал», ценовой запрос  "
        f"·  две отдельные заявки  ·  подача до {VALID_UNTIL}  ·  цены с НДС 22%"
    )
    ws["A2"].font = Font(name="Calibri", size=10, italic=True, color=WHITE)
    ws["A2"].fill = PatternFill("solid", fgColor=NAVY2)
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in range(1, 10):
        ws.cell(2, col).fill = PatternFill("solid", fgColor=NAVY2)
    ws.row_dimensions[2].height = 22

    meta = [
        ("Заказчик", "ФГКУ «Донской спасательный центр МЧС России»"),
        ("ИНН / КПП / ОГРН", "6102006605 / 610201001 / 1026100666525"),
        ("Адрес поставки", "346709, Ростовская обл., Аксайский р-н, п. Ковалевка, ул. Салютная, д. 2"),
        ("Контакт заказчика", "Пахомов И.В., donsc@dsc.61.mchs.gov.ru, +7 (86350) 2-71-98"),
        ("Поставщик", "к заполнению (DSSL / авторизованный дилер TRASSIR)"),
        ("Срок подачи на ЭТП", VALID_UNTIL),
        ("НМЦ", "не указана"),
        ("Оплата", "по 44-ФЗ и условиям контракта казённого учреждения"),
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

    nvr = enrich(ITEM_NVR)
    cam = enrich(ITEM_CAM)
    equiv = enrich(ITEM_EQUIV)
    anyip = enrich(ITEM_ANYIP)
    hdd = enrich(ITEM_HDD)
    nd = enrich(ITEM_ND)
    nvr_sum = nvr["sum"]
    cam_sum = cam["sum"]
    both_sum = money(nvr_sum + cam_sum)
    work_sum = money(both_sum + anyip["sum"])
    archive_sum = money(work_sum + hdd["sum"])
    full_sum = money(archive_sum + nd["sum"])

    section_bar(ws, 13, "А. Заявка 254364  ·  № ЭТП 998818  ·  NeuroStation Astra 9800R/128-S, 2 шт.")
    for col, h in enumerate(headers, 1):
        header_cell(ws, 14, col, h)
    write_item_row(ws, 15, nvr, GRAY)
    totals_block(ws, 16, "Итого по заявке 254364 (подавать отдельно)", nvr_sum, NAVY)

    section_bar(ws, 18, "Б. Заявка 254397  ·  № ЭТП 998881  ·  камера 4BHNN0-0-0-0, 51 шт.")
    for col, h in enumerate(headers, 1):
        header_cell(ws, 19, col, h)
    write_item_row(ws, 20, cam, WHITE)
    totals_block(ws, 21, "Итого по заявке 254397 (подавать отдельно)", cam_sum, NAVY)

    section_bar(ws, 23, "Справочно: сумма двух заявок (не объединять на ЭТП)")
    totals_block(ws, 24, "254364 + 254397", both_sum, "375623")

    section_bar(ws, 26, "Эквивалент камер — НЕ подавать, если документация не допускает замену")
    for col, h in enumerate(headers, 1):
        header_cell(ws, 27, col, h)
    write_item_row(ws, 28, equiv, GOLD)
    totals_block(ws, 29, "51× TR-D2151IR3 v2 (R) 2.8 (только при допуске эквивалента)", equiv["sum"], "C65911")

    section_bar(ws, 31, "В. Не входит в извещения (AnyIP, HDD — отдельная закупка / уточнение ТЗ)")
    for col, h in enumerate(headers, 1):
        header_cell(ws, 32, col, h)
    write_item_row(ws, 33, anyip, GOLD)
    write_item_row(ws, 34, hdd, GREEN)
    write_item_row(ws, 35, nd, ORANGE)

    section_bar(ws, 37, "Сводка (справочно)")
    summary_headers = ["Комплект", "Состав", "Сумма с НДС, ₽", "в т.ч. НДС 22%, ₽", "Без НДС, ₽"]
    for col, h in enumerate(summary_headers, 1):
        header_cell(ws, 38, col, h)
    ws.merge_cells("E38:I38")
    for col in range(5, 10):
        ws.cell(38, col).fill = PatternFill("solid", fgColor=NAVY)
        ws.cell(38, col).border = THIN

    packs = [
        ("Заявка 254364", "2× NeuroStation", nvr_sum, LIGHT),
        ("Заявка 254397", "51× 4BHNN0-0-0-0", cam_sum, LIGHT),
        ("Обе заявки", "NVR + камеры по извещениям", both_sum, GOLD),
        ("+ AnyIP Astra", "чтобы NVR писал камеры", work_sum, GREEN),
        ("+ HDD 10 ТБ ×8", "~30 суток архива", archive_sum, ORANGE),
        ("+ Neuro Detector", "аналитика на 51 канал", full_sum, GRAY),
    ]
    for i, (name, composition, gross, fill) in enumerate(packs):
        row = 39 + i
        ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=9)
        ws.cell(row, 1, name).font = Font(name="Calibri", size=10, bold=True)
        ws.cell(row, 2, composition).font = Font(name="Calibri", size=10)
        for col, val in ((3, gross), (4, vat_from_gross(gross)), (5, net_from_gross(gross))):
            c = ws.cell(row, col, val)
            c.number_format = '#,##0.00'
            c.font = Font(name="Calibri", size=10, bold=(i == 2))
            c.alignment = Alignment(horizontal="right", vertical="center")
        for col in range(1, 10):
            ws.cell(row, col).fill = PatternFill("solid", fgColor=fill)
            ws.cell(row, col).border = THIN
            ws.cell(row, col).alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[row].height = 20

    notes = [
        "Условия и оговорки (44-ФЗ)",
        "• Две независимые процедуры. На ЭТП ГПБ подавать две заявки до 18.08.2026, 00:00 МСК. НМЦ не указана.",
        "• Предмет камер — 51× 4BHNN0-0-0-0. Смесь 1 купол + 49 цилиндр + 1 PTZ предмету заявки не соответствует.",
        "• Эквивалент TRASSIR — только если документация прямо допускает. Иначе заявка может быть отклонена.",
        "• Артикул 4BHNN0 в открытых каталогах не найден; 22 000 ₽ — ориентир. Подтвердить у DSSL/НИЦ до подачи.",
        "• NVR без HDD и без лицензий AnyIP (Astra Linux, 95887). В извещениях этого нет — система «из коробки» не запишет 51 камеру.",
        "• Оплата — по 44-ФЗ / контракту ФГКУ, не коммерческая схема 70/30.",
        "• Поставка: п. Ковалевка, ул. Салютная, 2. Срок — по документации процедуры.",
        "• Контакты: заказчик Пахомов И.В. donsc@dsc.61.mchs.gov.ru; DSSL 8 (800) 100-91-12, код NVR 80222.",
    ]
    section_bar(ws, 46, notes[0])
    for i, text in enumerate(notes[1:], 47):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=9)
        cell = ws.cell(i, 1, text)
        cell.font = Font(name="Calibri", size=9, color=RED_NOTE if i in (47, 48, 49, 51) else "000000")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[i].height = 18

    ws.merge_cells("A56:D56")
    ws["A56"] = "Поставщик / участник закупки"
    ws["A56"].font = Font(name="Calibri", size=10, bold=True)
    ws.merge_cells("F56:I56")
    ws["F56"] = "Заказчик"
    ws["F56"].font = Font(name="Calibri", size=10, bold=True)
    ws.merge_cells("A58:D58")
    ws["A58"] = "________________ / ________________"
    ws.merge_cells("F58:I58")
    ws["F58"] = "________________ / ________________"
    ws.merge_cells("A59:D59")
    ws["A59"] = "подпись, ФИО, дата"
    ws["A59"].font = Font(name="Calibri", size=8, italic=True, color="808080")
    ws.merge_cells("F59:I59")
    ws["F59"] = "подпись, ФИО, дата"
    ws["F59"].font = Font(name="Calibri", size=8, italic=True, color="808080")

    ws.freeze_panes = "A13"
    ws.print_area = "A1:I59"
    ws.oddHeader.left.text = f"КП {KP_NO} ред. 4"
    ws.oddFooter.center.text = "Страница &P из &N · две заявки ЭТП ГПБ · цены с НДС 22%"

    return {
        "nvr": nvr_sum,
        "cam": cam_sum,
        "both": both_sum,
        "equiv": equiv["sum"],
        "work": work_sum,
        "archive": archive_sum,
        "full": full_sum,
        "nvr_vat": vat_from_gross(nvr_sum),
        "cam_vat": vat_from_gross(cam_sum),
        "both_vat": vat_from_gross(both_sum),
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
            "Камера по извещению 254397",
            "4BHNN0-0-0-0",
            22000,
            "артикул в открытых каталогах не найден; ориентир опта 4 Мп НИЦ/Nexus",
        ),
        (
            "Эквивалент TRASSIR СТ-1 (не подавать без допуска)",
            "TR-D2151IR3 v2 (R) 2.8",
            16800,
            "https://www.dssl.ru/products/tr-d2151ir3-v2-r-2-8-ip-kamera/",
        ),
        (
            "TRASSIR AnyIP (Astra Linux)",
            "95887",
            5090,
            "https://www.dssl.ru/products/litsenziya-trassir-anyip-astra-linux/ (проектная; ориентир AnyIP OS 57885)",
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
