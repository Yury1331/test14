#!/usr/bin/env python3
"""Ориентировочное КП: 2× NeuroStation Astra 9800R/128-S + 51× 4BHNN0-0-0-0."""

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
            "2 лицензии Neuro Detector в комплекте"
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
            "Камера видеонаблюдения цифровая 4 Мп, прошивка Nexus. "
            "Артикул 4BHNN0-0-0-0 (производитель/корпус/объектив — подтвердить у НИЦ)"
        ),
        "sku": "4BHNN0-0-0-0",
        "unit": "шт.",
        "qty": 51,
        "price": 22000.00,
        "note": "Ориентир опта НИЦ 4 Мп: 15–25 тыс. ₽; цена подлежит подтверждению",
    },
]

ITEM_ANYIP = {
    "no": 3,
    "name": "Лицензия TRASSIR AnyIP (TRASSIR OS) — подключение 1 IP-камеры любого производителя",
    "sku": "DSSL 57885",
    "unit": "шт.",
    "qty": 51,
    "price": 5090.00,
    "note": "Цена DSSL с 01.04.2026; без лицензии камера к NVR не подключается",
}

ITEM_HDD = {
    "no": 4,
    "name": "HDD 10 ТБ Seagate SkyHawk AI, 3.5\" SATA, рекомендован DSSL для NeuroStation 8-disk",
    "sku": "ST10000VE001",
    "unit": "шт.",
    "qty": 8,
    "price": 55200.00,
    "note": "Ориентир ANDPRO; ~30 суток архива при 51×4 Мп / 3 Мбит/с",
}

ITEM_ND = {
    "no": 5,
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
        f"№ {KP_NO} от {DATE}  ·  Поставка видеорегистраторов TRASSIR и цифровых камер  ·  "
        f"Цены с НДС 22%, ориентир открытого рынка, не оферта"
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
        ("Поставщик", "к заполнению (ориентир: дилеры TRASSIR / DSSL, НИЦ «Технологии»)"),
        ("Срок действия цен", VALID_UNTIL),
        ("Срок поставки", "NVR: 5–15 раб. дней при наличии; камеры НИЦ — под заказ"),
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
    write_item_row(ws, 13, request_items[0], GRAY)
    write_item_row(ws, 14, request_items[1], WHITE)
    request_sum = money(sum(x["sum"] for x in request_items))
    totals_block(ws, 15, "Итого по запросу (позиции 1–2)", request_sum, NAVY)

    ws.merge_cells("A17:I17")
    ws["A17"] = (
        "2. Обязательно к поставке для работоспособности "
        "(в регистраторе нет лицензий на IP-камеры; камеры не native TRASSIR)"
    )
    ws["A17"].font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    ws["A17"].fill = PatternFill("solid", fgColor=NAVY)
    for col in range(1, 10):
        ws.cell(17, col).fill = PatternFill("solid", fgColor=NAVY)
    ws.row_dimensions[17].height = 20

    for col, h in enumerate(headers, 1):
        header_cell(ws, 18, col, h)
    anyip = enrich(ITEM_ANYIP)
    write_item_row(ws, 19, anyip, GOLD)
    work_sum = money(request_sum + anyip["sum"])
    totals_block(ws, 20, "Итого работоспособный минимум (позиции 1–3)", work_sum, "375623")

    ws.merge_cells("A22:I22")
    ws["A22"] = "3. Рекомендуемые опции (в исходном запросе отсутствуют)"
    ws["A22"].font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    ws["A22"].fill = PatternFill("solid", fgColor=NAVY)
    for col in range(1, 10):
        ws.cell(22, col).fill = PatternFill("solid", fgColor=NAVY)

    for col, h in enumerate(headers, 1):
        header_cell(ws, 23, col, h)
    hdd = enrich(ITEM_HDD)
    nd = enrich(ITEM_ND)
    write_item_row(ws, 24, hdd, GREEN)
    write_item_row(ws, 25, nd, ORANGE)

    archive_sum = money(work_sum + hdd["sum"])
    full_sum = money(archive_sum + nd["sum"])
    totals_block(ws, 26, "Итого с архивом ~30 суток (позиции 1–4)", archive_sum, "548235")
    totals_block(ws, 27, "Итого полный комплект с аналитикой (позиции 1–5)", full_sum, NAVY)

    ws.merge_cells("A29:I29")
    ws["A29"] = "4. Сводка комплектов"
    ws["A29"].font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    ws["A29"].fill = PatternFill("solid", fgColor=NAVY)
    for col in range(1, 10):
        ws.cell(29, col).fill = PatternFill("solid", fgColor=NAVY)

    summary_headers = ["Комплект", "Состав", "Сумма с НДС, ₽", "в т.ч. НДС 22%, ₽", "Без НДС, ₽"]
    for col, h in enumerate(summary_headers, 1):
        header_cell(ws, 30, col, h)
    ws.merge_cells("E30:I30")
    for col in range(5, 10):
        ws.cell(30, col).fill = PatternFill("solid", fgColor=NAVY)
        ws.cell(30, col).border = THIN

    packs = [
        ("А. Только запрошенные позиции", "2 NVR + 51 камера", request_sum, LIGHT),
        ("Б. Работоспособный минимум", "А + 51× AnyIP", work_sum, GOLD),
        ("В. С архивом 30 суток", "Б + 8× HDD 10 ТБ", archive_sum, GREEN),
        ("Г. Полный комплект", "В + Neuro Detector на все камеры", full_sum, ORANGE),
    ]
    for i, (name, composition, gross, fill) in enumerate(packs):
        row = 31 + i
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
        "• Документ — ориентировочный расчёт по открытым ценам на 14.08.2026, не является офертой. Фиксация — в КП DSSL/дилера TRASSIR и НИЦ.",
        "• Регистратор поставляется без HDD и без лицензий на подключение IP-камер. В каждом NVR — 2 лицензии Neuro Detector.",
        "• Артикул камеры 4BHNN0-0-0-0 в публичных каталогах не найден; по коду — 4 Мп НИЦ «Технологии», прошивка Nexus, базовая комплектация. Корпус и объектив подтвердить до заказа.",
        "• Два NVR на 51 камеру: запас каналов 256 vs 51 (резерв, два объекта или расширение). Запись можно вести на одном регистраторе.",
        "• Оценка архива: 51×4 Мп, H.265, 3 Мбит/с ≈ 1,65 ТБ/сутки; 8×10 ТБ достаточно примерно на 30 суток на одном NVR.",
        "• Не включено: PoE-коммутаторы, СКС, шкаф 19\", ИБП, мониторы, монтаж, ПНР, доставка.",
        "• Гарантия TRASSIR — по паспорту (типично 3–5 лет на NVR). Камеры НИЦ — уточняется (часто 3 года).",
        "• Контакты: DSSL 8 (800) 100-91-12, dssl.ru, код NVR 80222; НИЦ nic-tech.ru, 8 (800) 555-47-65.",
    ]
    ws.merge_cells("A36:I36")
    ws["A36"] = notes[0]
    ws["A36"].font = Font(name="Calibri", size=12, bold=True, color=WHITE)
    ws["A36"].fill = PatternFill("solid", fgColor=NAVY)
    for col in range(1, 10):
        ws.cell(36, col).fill = PatternFill("solid", fgColor=NAVY)

    for i, text in enumerate(notes[1:], 37):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=9)
        cell = ws.cell(i, 1, text)
        cell.font = Font(name="Calibri", size=9, color=RED_NOTE if i == 39 else "000000")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[i].height = 18

    ws.merge_cells("A46:D46")
    ws["A46"] = "Поставщик / исполнитель"
    ws["A46"].font = Font(name="Calibri", size=10, bold=True)
    ws.merge_cells("F46:I46")
    ws["F46"] = "Заказчик"
    ws["F46"].font = Font(name="Calibri", size=10, bold=True)
    ws.merge_cells("A48:D48")
    ws["A48"] = "________________ / ________________"
    ws.merge_cells("F48:I48")
    ws["F48"] = "________________ / ________________"
    ws.merge_cells("A49:D49")
    ws["A49"] = "подпись, ФИО, дата"
    ws["A49"].font = Font(name="Calibri", size=8, italic=True, color="808080")
    ws.merge_cells("F49:I49")
    ws["F49"] = "подпись, ФИО, дата"
    ws["F49"].font = Font(name="Calibri", size=8, italic=True, color="808080")

    ws.auto_filter.ref = "A12:I14"
    ws.freeze_panes = "A13"
    ws.print_area = "A1:I49"
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
            "Камера 4 Мп НИЦ (ориентир опта)",
            "4BHNN0-0-0-0 / NIC-4-BUL-Moto-RUS",
            22000,
            "https://tonzar.com — диапазон 15 000–25 000 ₽, в КП середина 22 000 ₽",
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
