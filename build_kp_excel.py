#!/usr/bin/env python3
"""Generate Excel price comparison with VAT for all suppliers."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

VAT_RATE = 0.20

# DSSL: explicit gross (with VAT) and VAT from KP ДП-00069239
DSSL = {
    "doc": "КП № ДП-00069239 от 15.07.2026",
    "valid_until": "17.07.2026",
    "delivery": "3–5 раб. дней до склада поставщика",
    "items": [
        {
            "no": 1,
            "name": "IP-видеокамера уличная цилиндрическая",
            "model_dssl": "DH-IPC-HFW1439TL1P-A-IL-0280B",
            "model_itv": "DH-IPC-HFW2449SP-S-IL-0280B",
            "qty": 5,
            "unit": "шт",
            "dssl_price": 5034.00,
            "dssl_sum": 25170.00,
            "dssl_vat": 4538.85,
            "itv_price": 5093.00,
            "itv_sum": 25465.00,
            "comment": "ITV: указана как аналог DH-IPC-HFW1439TL1P-A-IL-0280B",
        },
        {
            "no": 2,
            "name": "PoE-коммутатор 8 портов",
            "model_dssl": "DH-CS4010-8ET2GT-110",
            "model_itv": "DH-CS4010-8ET2GT-110",
            "qty": 1,
            "unit": "шт",
            "dssl_price": 5394.00,
            "dssl_sum": 5394.00,
            "dssl_vat": 972.69,
            "itv_price": 3725.00,
            "itv_sum": 3725.00,
            "comment": "",
        },
        {
            "no": 3,
            "name": "IP-видеорегистратор 8 каналов 4K",
            "model_dssl": "DHI-NVR2108HS-4KS3",
            "model_itv": "DHI-NVR2108HS-4KS3",
            "qty": 1,
            "unit": "шт",
            "dssl_price": 7254.00,
            "dssl_sum": 7254.00,
            "dssl_vat": 1308.10,
            "itv_price": 4680.00,
            "itv_sum": 4680.00,
            "comment": "",
        },
        {
            "no": 4,
            "name": "Жёсткий диск 8 Тб Seagate SkyHawk AI",
            "model_dssl": "ST8000VE001",
            "model_itv": None,
            "qty": 1,
            "unit": "шт",
            "dssl_price": 43600.00,
            "dssl_sum": 43600.00,
            "dssl_vat": 7862.30,
            "itv_price": None,
            "itv_sum": None,
            "comment": "В оферте ITV GROUP не указан",
        },
        {
            "no": 5,
            "name": "Доставка",
            "model_dssl": "—",
            "model_itv": None,
            "qty": 1,
            "unit": "усл.",
            "dssl_price": 1500.00,
            "dssl_sum": 1500.00,
            "dssl_vat": 270.49,
            "itv_price": None,
            "itv_sum": None,
            "comment": "В оферте ITV GROUP не указана",
        },
    ],
}


def vat_from_gross(gross: float) -> float:
    return round(gross * VAT_RATE / (1 + VAT_RATE), 2)


def net_from_gross(gross: float) -> float:
    return round(gross - vat_from_gross(gross), 2)


def style_header(cell, fill_color="1F4E79"):
    cell.font = Font(bold=True, color="FFFFFF", size=10)
    cell.fill = PatternFill("solid", fgColor=fill_color)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def style_subheader(cell, fill_color="D6E4F0"):
    cell.font = Font(bold=True, size=10)
    cell.fill = PatternFill("solid", fgColor=fill_color)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def apply_border(ws, min_row, max_row, min_col, max_col):
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.border = border


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "Сравнение цен с НДС"

    ws.merge_cells("A1:N1")
    ws["A1"] = (
        "Сравнение цен по поставщикам (с НДС) — МБУ «Сочисвет», "
        "система охранного видеонаблюдения"
    )
    ws["A1"].font = Font(bold=True, size=12)
    ws["A1"].alignment = Alignment(horizontal="center")

    headers_row1 = [
        ("A2", "№"),
        ("B2", "Наименование"),
        ("C2", "Кол-во"),
        ("D2", "Ед."),
        ("E2", 'ООО «ДССЛ-Первый»'),
        ("I2", "ITV GROUP / IPDROM"),
        ("M2", 'ООО «Сфера-95»'),
        ("N2", "Примечание"),
    ]
    for addr, val in headers_row1:
        ws[addr] = val
        style_header(ws[addr])

    ws.merge_cells("E2:H2")
    ws.merge_cells("I2:L2")
    ws["M2"] = 'ООО «Сфера-95»'
    style_header(ws["M2"], "5B5B5B")
    style_header(ws["N2"], "5B5B5B")

    subheaders = [
        "Артикул / модель",
        "Цена с НДС, ₽",
        "Сумма с НДС, ₽",
        "НДС, ₽",
    ]
    for i, title in enumerate(subheaders):
        col = 5 + i
        cell = ws.cell(row=3, column=col, value=title)
        style_subheader(cell)
        cell = ws.cell(row=3, column=9 + i, value=title)
        style_subheader(cell, "E2EFDA")

    ws["M3"] = "Цена с НДС, ₽"
    ws["N3"] = "Примечание"
    style_subheader(ws["M3"], "EDEDED")
    style_subheader(ws["N3"], "EDEDED")

    for col in (1, 2, 3, 4):
        cell = ws.cell(row=3, column=col, value=ws.cell(row=2, column=col).value)
        style_subheader(cell)
    ws.merge_cells("A2:A3")
    ws.merge_cells("B2:B3")
    ws.merge_cells("C2:C3")
    ws.merge_cells("D2:D3")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["C2"].alignment = Alignment(horizontal="center", vertical="center")
    ws["D2"].alignment = Alignment(horizontal="center", vertical="center")

    row = 4
    dssl_total = dssl_vat_total = 0.0
    itv_total = itv_vat_total = 0.0

    for item in DSSL["items"]:
        ws.cell(row=row, column=1, value=item["no"])
        ws.cell(row=row, column=2, value=item["name"])
        ws.cell(row=row, column=3, value=item["qty"])
        ws.cell(row=row, column=4, value=item["unit"])

        ws.cell(row=row, column=5, value=item["model_dssl"])
        ws.cell(row=row, column=6, value=item["dssl_price"]).number_format = "#,##0.00"
        ws.cell(row=row, column=7, value=item["dssl_sum"]).number_format = "#,##0.00"
        ws.cell(row=row, column=8, value=item["dssl_vat"]).number_format = "#,##0.00"
        dssl_total += item["dssl_sum"]
        dssl_vat_total += item["dssl_vat"]

        if item["itv_price"] is not None:
            itv_vat = vat_from_gross(item["itv_sum"])
            ws.cell(row=row, column=9, value=item["model_itv"])
            ws.cell(row=row, column=10, value=item["itv_price"]).number_format = "#,##0.00"
            ws.cell(row=row, column=11, value=item["itv_sum"]).number_format = "#,##0.00"
            ws.cell(row=row, column=12, value=itv_vat).number_format = "#,##0.00"
            itv_total += item["itv_sum"]
            itv_vat_total += itv_vat
        else:
            for c in range(9, 13):
                ws.cell(row=row, column=c, value="—")

        ws.cell(row=row, column=13, value="нет КП")
        note = item["comment"]
        if item["itv_sum"] is not None and item["itv_price"] is not None:
            note = (note + "; " if note else "") + "НДС ITV: не указан в оферте, рассчитан 20%"
        ws.cell(row=row, column=14, value=note)

        for col in range(1, 15):
            ws.cell(row=row, column=col).alignment = Alignment(
                vertical="center", wrap_text=True
            )
            if col in (1, 3, 4):
                ws.cell(row=row, column=col).alignment = Alignment(
                    horizontal="center", vertical="center"
                )
        row += 1

    # Totals
    ws.cell(row=row, column=2, value="ИТОГО").font = Font(bold=True)
    ws.cell(row=row, column=7, value=dssl_total).number_format = "#,##0.00"
    ws.cell(row=row, column=8, value=dssl_vat_total).number_format = "#,##0.00"
    ws.cell(row=row, column=11, value=itv_total).number_format = "#,##0.00"
    ws.cell(row=row, column=12, value=itv_vat_total).number_format = "#,##0.00"
    for col in (2, 7, 8, 11, 12):
        ws.cell(row=row, column=col).font = Font(bold=True)
    row += 1

    ws.cell(row=row, column=2, value="Сопоставимый комплект (поз. 1–3)")
    ws.cell(row=row, column=7, value=37818.00).number_format = "#,##0.00"
    ws.cell(row=row, column=8, value=6819.64).number_format = "#,##0.00"
    ws.cell(row=row, column=11, value=33870.00).number_format = "#,##0.00"
    ws.cell(row=row, column=12, value=5645.00).number_format = "#,##0.00"
    row += 2

    info = [
        ("Документ ДССЛ-Первый:", DSSL["doc"]),
        ("Срок действия цен ДССЛ:", DSSL["valid_until"]),
        ("Срок поставки ДССЛ:", DSSL["delivery"]),
        ("Документ ITV GROUP:", "Письмо-оферта (без номера КП)"),
        ("Срок поставки ITV:", "3 раб. дня после оплаты (склад Москва)"),
        ("Документ Сфера-95:", "Акт обследования и ППИ — цены отсутствуют"),
        (
            "Примечание по НДС ITV:",
            "В оферте ставка НДС не указана; для сравнения принято, что цены включают НДС 20%, "
            "сумма НДС рассчитана как 20/120 от суммы.",
        ),
    ]
    for label, text in info:
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=14)
        ws.cell(row=row, column=2, value=text).alignment = Alignment(wrap_text=True)
        row += 1

    apply_border(ws, 2, row - len(info) - 2, 1, 14)

    widths = {
        "A": 5,
        "B": 34,
        "C": 8,
        "D": 6,
        "E": 28,
        "F": 14,
        "G": 16,
        "H": 12,
        "I": 28,
        "J": 14,
        "K": 16,
        "L": 12,
        "M": 14,
        "N": 36,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 28
    ws.row_dimensions[3].height = 36
    ws.freeze_panes = "A4"

    out = "/workspace/Сравнение_КП_цены_с_НДС.xlsx"
    wb.save(out)
    print(out)


if __name__ == "__main__":
    main()
