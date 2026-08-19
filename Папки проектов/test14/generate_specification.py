#!/usr/bin/env python3
"""Generate 03_Спецификация.xlsx with extended equipment descriptions."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ITEMS = [
    {
        "name": "Видеокамера DS-2CD2043G2-IU(2.8mm)",
        "model": "DS-2CD2043G2-IU(2.8mm)",
        "qty": 5,
        "unit_price": 7390,
        "description_extended": (
            "Hikvision, IP-камера цилиндрическая (bullet), серия AcuSense. "
            "Разрешение 4 Мп (2688×1520), матрица 1/3\" Progressive Scan CMOS. "
            "Объектив 2,8 мм (фикс.), угол обзора ~105°. "
            "Подсветка Smart IR до 30 м, WDR 120 dB, H.265+/H.264+, PoE (802.3af). "
            "Аналитика: классификация человек/транспорт, детекция вторжения в зону, "
            "пересечение линии, обнаружение лиц (базовое). "
            "Корпус IP67, рабочая t −30…+60 °C. "
            "Назначение: уличные/внутренние точки наблюдения узла доступа, цилиндрический форм-фактор по ТЗ."
        ),
    },
    {
        "name": "Видеорегистратор DS-7608NXI-K2/8P",
        "model": "DS-7608NXI-K2/8P(D)",
        "qty": 1,
        "unit_price": 16270,
        "description_extended": (
            "Hikvision, сетевой видеорегистратор 8 каналов с 8 портами PoE (K-серия). "
            "Входящая пропускная способность до 80 Мбит/с, разрешение записи до 4K. "
            "Поддержка H.265+/H.264+, 1×HDD до 10 ТБ (HDD в комплекте отдельной строкой). "
            "Встроенный PoE-коммутатор 802.3af/at на 8 портов — питание камер узла без отдельных БП. "
            "HDMI/VGA, удалённый доступ через клиент/браузер, ONVIF. "
            "В спецификации заказчика: DS-7608NXI-K2/8P; в КП поставщика указана модификация DS-7608NXI-K2/8P(D) — "
            "уточнить эквивалентность и комплектацию перед публикацией ТЗ."
        ),
    },
    {
        "name": "Видеокамера IP Hikvision DS-2DE5432IW-AE(S5) 4Мп уличная",
        "model": "DS-2DE5432IW-AE(T5)",
        "qty": 1,
        "unit_price": 42890,
        "description_extended": (
            "Hikvision, IP PTZ-камера уличная, 4 Мп, серия Pro. "
            "Оптический зум 32×, скорость поворота до 100°/с, предустановки/патрулирование. "
            "Smart IR до 150 м, WDR, H.265+, PoE+. "
            "Защита IP66, t −40…+70 °C — для обзора перекрёстка/периметра узла. "
            "Расхождение черновика и КП: в ТЗ указана модификация (S5), в прайсе поставщика — (T5). "
            "Требуется сверка ревизии, наличия IR-прожектора и совместимости с NVR DS-7608NXI-K2/8P."
        ),
    },
    {
        "name": "Коммутатор PoE Dahua DH-CS4010-8ET2GT-110 управляемый",
        "model": "DH-CS4010-8ET2GT-110",
        "qty": 1,
        "unit_price": 4230,
        "description_extended": (
            "Dahua, управляемый PoE-коммутатор: 8 портов 10/100 Мбит/с PoE + 2 порта Gigabit uplink. "
            "Бюджет PoE до 110 Вт, стандарты 802.3af/at. "
            "Управление через web/CLI, VLAN, QoS, loop detection. "
            "Назначение: агрегация камер узла и uplink к NVR/локальной сети при недостатке PoE-портов регистратора "
            "или для резервирования топологии. Модель совпадает в ТЗ и КП."
        ),
    },
    {
        "name": "Жесткий диск ST2000VX017",
        "model": "ST2000VX017",
        "qty": 1,
        "unit_price": 15570,
        "description_extended": (
            "Seagate SkyHawk, HDD 2 ТБ, 3,5\", 5900 об/мин, 256 МБ кэш, SATA 6 Гбит/с. "
            "Линейка для систем видеонаблюдения (до 64 камер / 32 потока), "
            "оптимизация записи 24/7, workload rate ~180 ТБ/год. "
            "Устанавливается в NVR DS-7608NXI-K2/8P для архива узла доступа. "
            "Модель совпадает в ТЗ и КП."
        ),
    },
]

HEADERS = [
    "№",
    "Наименование",
    "Модель (артикул)",
    "Кол-во",
    "Цена за ед., ₽",
    "Сумма, ₽",
    "Расширенное описание",
]


def build_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = "Спецификация"

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)

    ws.append(HEADERS)
    for col in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

    total = 0
    for i, item in enumerate(ITEMS, start=1):
        line_total = item["qty"] * item["unit_price"]
        total += line_total
        ws.append([
            i,
            item["name"],
            item["model"],
            item["qty"],
            item["unit_price"],
            line_total,
            item["description_extended"],
        ])

    total_row = ws.max_row + 1
    ws.cell(row=total_row, column=5, value="Итого, ₽:")
    ws.cell(row=total_row, column=5).font = Font(bold=True)
    ws.cell(row=total_row, column=6, value=total)
    ws.cell(row=total_row, column=6).font = Font(bold=True)

    widths = [5, 42, 28, 8, 14, 14, 72]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{ws.max_row - 1}"
    return wb


if __name__ == "__main__":
    out = "/workspace/Папки проектов/test14/03_Спецификация.xlsx"
    build_workbook().save(out)
    print(f"Saved: {out} ({len(ITEMS)} items)")
