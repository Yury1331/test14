#!/usr/bin/env python3
"""Generate Excel registry of video surveillance tenders for Sanatorium Raduga (INN 2320095012)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = "https://zakupki.gov.ru/epz/order/notice"
SEARCH = "https://zakupki.gov.ru/epz/order/extendedsearch/results.html?searchString="

CUSTOMER = 'ФБЛПУ «Санаторий «Радуга» ФНС России»'
INN = "2320095012"
CONTACT = "omts@sochi-raduga.ru"


def link(notice_type: str, reg: str) -> str:
    return f"{BASE}/{notice_type}/view/common-info.html?regNumber={reg}"


LOTS = [
    {
        "category": "Видеонаблюдение",
        "reg": "337178215",
        "type": None,
        "procedure": "Анализ цен (вне ЕИС)",
        "subject": "Поставка и монтаж камер видеонаблюдения, 20 шт. (КТРУ 26.70.13.000)",
        "nmck": None,
        "contract_price": None,
        "status": "Приём КП завершён 21.07.2026; ожидается 44-ФЗ",
        "platform": "— (КП на e-mail)",
        "deadline": "21.07.2026",
        "published": "17.07.2026",
        "external_url": "https://zakupki360.ru/tender/97360330",
        "note": "Номер 0318100043126000013 в ЕИС — другая закупка (ТО автобуса)",
    },
    {
        "category": "Инфраструктура ВН / ИТ",
        "reg": "0318100043126000023",
        "type": "ea20",
        "procedure": "Электронный аукцион",
        "subject": "Поставка серверов: 1×2 млн + 2×1,26 млн ₽ (3 шт., КТРУ 26.20.14.000)",
        "nmck": 4513333,
        "contract_price": 2835418,
        "status": "Приём завершён 11.06.2026; работа комиссии (2 заявки)",
        "platform": "Сбербанк-АСТ",
        "deadline": "11.06.2026",
        "published": "28.05.2026",
        "external_url": "https://zakupkipro.com/zakupka/0318100043126000023/",
        "note": "Мин. заявка 2 835 418 ₽",
    },
    {
        "category": "Проектирование (смежное)",
        "reg": "0318100043126000029",
        "type": "ea20",
        "procedure": "Электронный аукцион",
        "subject": "Проектирование капремонта клуб-столовой с водолечебницей (литер В), зона регистрации посетителей, ул. Виноградная, 53 стр.8",
        "nmck": 5354700,
        "contract_price": None,
        "status": "Приём завершён 06.2026",
        "platform": "Сбербанк-АСТ",
        "deadline": "06.06.2026",
        "published": "29.05.2026",
        "external_url": "https://zakupkipro.com/zakazchik/2320095012/",
        "note": "Проект может включать разделы ОС/СКУД/ВН",
    },
    {
        "category": "Видеонаблюдение / мультимедиа",
        "reg": "0318100043125000007",
        "type": "ea20",
        "procedure": "Электронный аукцион",
        "subject": "Поставка и установка системы мультимедиа (PTZ-видеокамера, IP-камера, акустика, усилители)",
        "nmck": 7033481,
        "contract_price": 7033481,
        "status": "Контракт заключён 04.2025",
        "platform": "Сбербанк-АСТ",
        "deadline": "04.2025",
        "published": "27.03.2025",
        "external_url": "https://www.tenderguru.ru/tender/85017717",
        "note": "Закупка у СМП",
    },
    {
        "category": "Безопасность (не ВН)",
        "reg": "0318100043125000065",
        "type": "ea20",
        "procedure": "Электронный аукцион",
        "subject": "ТО ОПС, СОУЭ, дымоудаления, пожаротушения (11 мес.)",
        "nmck": 726000,
        "contract_price": 722370,
        "status": "Контракт №3-1 от 26.01.2026; ООО «КОБРА ГАРАНТ СОЧИ»",
        "platform": "Сбербанк-АСТ",
        "deadline": "12.01.2026",
        "published": "29.12.2025",
        "external_url": "https://synapsenet.ru/zakupki/fz44/0318100043125000065",
        "note": "Исполнение до 24.02.2027",
    },
]

HEADERS = [
    "№",
    "Категория",
    "№ закупки / ID",
    "Способ определения",
    "Заказчик",
    "Предмет закупки",
    "НМЦК, ₽",
    "Цена контракта, ₽",
    "Статус",
    "ЭТП",
    "Размещено",
    "Окончание заявок",
    "Контакт",
    "Ссылка ЕИС",
    "Поиск в ЕИС",
    "Доп. источник",
    "Примечание",
]


def fmt_money(value):
    if value is None:
        return "—"
    return value


def build_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = "Тендеры ВН"

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)

    ws.append(HEADERS)
    for col in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

    for i, lot in enumerate(LOTS, start=1):
        reg = lot["reg"]
        url = link(lot["type"], reg) if lot["type"] else lot["external_url"]
        search_url = SEARCH + (reg if lot["type"] else INN)
        row = [
            i,
            lot["category"],
            reg,
            lot["procedure"],
            CUSTOMER,
            lot["subject"],
            fmt_money(lot["nmck"]),
            fmt_money(lot["contract_price"]),
            lot["status"],
            lot["platform"],
            lot["published"],
            lot["deadline"],
            CONTACT,
            url if lot["type"] else "—",
            search_url,
            lot["external_url"],
            lot["note"],
        ]
        ws.append(row)
        r = ws.max_row
        if lot["type"]:
            link_cell = ws.cell(row=r, column=14)
            link_cell.hyperlink = url
            link_cell.value = url
            link_cell.font = Font(color="0563C1", underline="single")
        search_cell = ws.cell(row=r, column=15)
        search_cell.hyperlink = search_url
        search_cell.value = search_url
        search_cell.font = Font(color="0563C1", underline="single")
        ext_cell = ws.cell(row=r, column=16)
        ext_cell.hyperlink = lot["external_url"]
        ext_cell.value = lot["external_url"]
        ext_cell.font = Font(color="0563C1", underline="single")

    widths = [5, 22, 22, 22, 34, 48, 14, 14, 30, 16, 12, 14, 22, 50, 50, 50, 36]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return wb


if __name__ == "__main__":
    out = "/workspace/Санаторий Радуга/реестр-тендеров-видеонаблюдение.xlsx"
    build_workbook().save(out)
    print(f"Saved: {out} ({len(LOTS)} lots)")
