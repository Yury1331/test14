#!/usr/bin/env python3
"""Generate Excel table with ENS and OKPD2 codes for project equipment."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

EQUIPMENT = [
    {
        "no": 1,
        "name": "IP-видеокамера уличная цилиндрическая 4 Мп, 2,8 мм",
        "model": "DH-IPC-HFW1439TL1P-A-IL-0280B",
        "model_alt": "DH-IPC-HFW2449SP-S-IL-0280B (аналог ITV)",
        "qty": 5,
        "unit": "шт",
        "ens": "—",
        "okpd2": "26.40.33.111",
        "okpd2_name": "Видеокамеры для систем видеонаблюдения",
        "source": "ДССЛ / ITV / ППИ Сфера-95",
    },
    {
        "no": 2,
        "name": "PoE-коммутатор 8 портов",
        "model": "DH-CS4010-8ET2GT-110",
        "model_alt": "—",
        "qty": 1,
        "unit": "шт",
        "ens": "—",
        "okpd2": "26.30.11.122",
        "okpd2_name": "Оборудование коммутации и маршрутизации",
        "source": "ДССЛ / ITV / ППИ Сфера-95",
    },
    {
        "no": 3,
        "name": "IP-видеорегистратор 8 каналов 4K",
        "model": "DHI-NVR2108HS-4KS3",
        "model_alt": "—",
        "qty": 1,
        "unit": "шт",
        "ens": "—",
        "okpd2": "26.40.33.114",
        "okpd2_name": "Видеорегистраторы",
        "source": "ДССЛ / ITV / ППИ Сфера-95",
    },
    {
        "no": 4,
        "name": "Жёсткий диск 8 Тб для видеонаблюдения",
        "model": "Seagate SkyHawk AI ST8000VE001",
        "model_alt": "—",
        "qty": 1,
        "unit": "шт",
        "ens": "—",
        "okpd2": "26.20.21.110",
        "okpd2_name": "Устройства запоминающие внутренние",
        "source": "ДССЛ",
    },
    {
        "no": 5,
        "name": "ИБП (автономия ≥ 3 суток)",
        "model": "по ТЗ, номинал не указан",
        "model_alt": "—",
        "qty": 1,
        "unit": "шт",
        "ens": "—",
        "okpd2": "27.11.40",
        "okpd2_name": "Источники бесперебойного питания (уточнить)",
        "source": "ППИ Сфера-95",
    },
    {
        "no": 6,
        "name": "Шкаф антивандальный 4–6U",
        "model": "не указан",
        "model_alt": "—",
        "qty": 1,
        "unit": "шт",
        "ens": "—",
        "okpd2": "27.12.40",
        "okpd2_name": "Шкафы, стойки для оборудования (уточнить)",
        "source": "ППИ Сфера-95",
    },
    {
        "no": 7,
        "name": "Кабель UTP cat.5e outdoor",
        "model": "—",
        "model_alt": "—",
        "qty": 200,
        "unit": "м",
        "ens": "—",
        "okpd2": "27.32.13",
        "okpd2_name": "Провода и кабели электронные (уточнить)",
        "source": "ППИ Сфера-95",
    },
    {
        "no": 8,
        "name": "Коробка монтажная для уличной камеры",
        "model": "—",
        "model_alt": "—",
        "qty": 5,
        "unit": "шт",
        "ens": "—",
        "okpd2": "27.33.13",
        "okpd2_name": "Аксессуары для электрооборудования (уточнить)",
        "source": "ППИ Сфера-95",
    },
    {
        "no": 9,
        "name": "Кронштейн выносной для уличной камеры",
        "model": "—",
        "model_alt": "—",
        "qty": 1,
        "unit": "шт",
        "ens": "—",
        "okpd2": "27.33.13",
        "okpd2_name": "Аксессуары для электрооборудования (уточнить)",
        "source": "ППИ Сфера-95",
    },
    {
        "no": 10,
        "name": "Труба ПВХ гофрированная d25–32",
        "model": "—",
        "model_alt": "—",
        "qty": 200,
        "unit": "м",
        "ens": "—",
        "okpd2": "22.21.29",
        "okpd2_name": "Изделия из пластмасс прочие (уточнить)",
        "source": "ППИ Сфера-95",
    },
    {
        "no": 11,
        "name": "Кабель питания ПВС 3×2,5",
        "model": "—",
        "model_alt": "—",
        "qty": 10,
        "unit": "м",
        "ens": "—",
        "okpd2": "27.32.13",
        "okpd2_name": "Провода и кабели электронные (уточнить)",
        "source": "ППИ Сфера-95",
    },
    {
        "no": 12,
        "name": "Блок розеток 19\"",
        "model": "—",
        "model_alt": "—",
        "qty": 1,
        "unit": "шт",
        "ens": "—",
        "okpd2": "27.33.13",
        "okpd2_name": "Аксессуары для электрооборудования (уточнить)",
        "source": "ППИ Сфера-95",
    },
]


def style_header(cell):
    cell.font = Font(bold=True, color="FFFFFF", size=10)
    cell.fill = PatternFill("solid", fgColor="1F4E79")
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
    ws.title = "Коды ЕНС и ОКПД2"

    ws.merge_cells("A1:J1")
    ws["A1"] = (
        "Коды ЕНС и ОКПД2 — оборудование МБУ «Сочисвет», "
        "система охранного видеонаблюдения"
    )
    ws["A1"].font = Font(bold=True, size=12)
    ws["A1"].alignment = Alignment(horizontal="center")

    headers = [
        "№",
        "Наименование",
        "Модель / артикул (основная)",
        "Альтернативная модель",
        "Кол-во",
        "Ед.",
        "Код ЕНС (МТС)",
        "Код ОКПД2",
        "Расшифровка ОКПД2",
        "Источник",
    ]
    for col, title in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col, value=title)
        style_header(cell)

    row = 3
    for item in EQUIPMENT:
        values = [
            item["no"],
            item["name"],
            item["model"],
            item["model_alt"],
            item["qty"],
            item["unit"],
            item["ens"],
            item["okpd2"],
            item["okpd2_name"],
            item["source"],
        ]
        for col, value in enumerate(values, start=1):
            cell = ws.cell(row=row, column=col, value=value)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if col in (1, 5, 6, 7, 8):
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        row += 1

    notes = [
        "",
        "ВАЖНО: коды ЕНС — внутренние номенклатурные коды ПАО «МТС» (формат XXX.XXX.XXXXXX).",
        "В загруженных КП и ППИ коды ЕНС не указаны.",
        "Для присвоения/подбора кода ЕНС необходимо обратиться в справочник номенклатуры МТС (SAP / Sourcing)",
        "или к ответственному за закупку / ведущему инженеру филиала.",
        "",
        "Коды ОКПД2 приведены как ориентир для закупочной документации; финальный код уточняется",
        "по характеристикам позиции и требованиям закупочной службы МТС.",
    ]
    for note in notes:
        ws.cell(row=row, column=1, value=note)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
        ws.cell(row=row, column=1).alignment = Alignment(wrap_text=True)
        if note.startswith("ВАЖНО"):
            ws.cell(row=row, column=1).font = Font(bold=True, color="C00000")
        row += 1

    apply_border(ws, 2, row - len(notes) - 1, 1, 10)

    widths = {
        "A": 5,
        "B": 34,
        "C": 30,
        "D": 28,
        "E": 8,
        "F": 6,
        "G": 18,
        "H": 14,
        "I": 36,
        "J": 22,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 36
    ws.freeze_panes = "A3"

    out = "/workspace/Коды_ЕНС_оборудование_Сочисвет.xlsx"
    wb.save(out)
    print(out)


if __name__ == "__main__":
    main()
