from datetime import datetime
from pathlib import Path
import sys

import pandas as pd

from booking_logic import find_booking_conflicts


INPUT_FILE = Path("input/car_bookings.xlsx")
OUTPUT_REPORT = Path("output/booking_result.txt")

COLUMNS = [
    "id",
    "car_name",
    "car_number",
    "employee_name",
    "department",
    "date",
    "start_time",
    "end_time",
    "route",
    "purpose",
    "status",
    "comment",
]


def configure_console_encoding() -> None:
    """Настраивает UTF-8 для русского текста в Windows-консоли."""
    if hasattr(sys.stdin, "reconfigure"):
        sys.stdin.reconfigure(encoding="utf-8")
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")


def read_bookings() -> pd.DataFrame:
    """Читает Excel-файл с бронями."""
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"Файл {INPUT_FILE} не найден.")

    bookings = pd.read_excel(INPUT_FILE, dtype=str).fillna("")
    return bookings.reindex(columns=COLUMNS, fill_value="")


def save_bookings(bookings: pd.DataFrame) -> None:
    """Сохраняет обновленный Excel-файл."""
    bookings.to_excel(INPUT_FILE, index=False)


def ask_required(prompt: str) -> str:
    """Запрашивает обязательное поле."""
    while True:
        value = input(prompt).strip()
        if value:
            return value
        print("Поле обязательно для заполнения.")


def ask_date() -> str:
    """Запрашивает дату в формате ГГГГ-ММ-ДД."""
    while True:
        value = ask_required("Дата поездки (ГГГГ-ММ-ДД): ")
        try:
            datetime.strptime(value, "%Y-%m-%d")
            return value
        except ValueError:
            print("Неверный формат даты. Пример: 2026-06-10")


def ask_time(prompt: str) -> str:
    """Запрашивает время в формате ЧЧ:ММ."""
    while True:
        value = ask_required(prompt)
        try:
            datetime.strptime(value, "%H:%M")
            return value
        except ValueError:
            print("Неверный формат времени. Пример: 09:30")


def ask_time_interval() -> tuple[str, str]:
    """Запрашивает и проверяет интервал времени."""
    while True:
        start_time = ask_time("Время начала (ЧЧ:ММ): ")
        end_time = ask_time("Время окончания (ЧЧ:ММ): ")
        start = datetime.strptime(start_time, "%H:%M")
        end = datetime.strptime(end_time, "%H:%M")

        if start < end:
            return start_time, end_time

        print("Время окончания должно быть позже времени начала.")


def show_cars(bookings: pd.DataFrame) -> pd.DataFrame:
    """Показывает список автомобилей из файла броней."""
    cars = bookings[["car_name", "car_number"]].drop_duplicates().reset_index(drop=True)

    print("\nДоступные автомобили:")
    for index, car in cars.iterrows():
        print(f"{index + 1}. {car['car_name']} ({car['car_number']})")

    return cars


def ask_car(cars: pd.DataFrame) -> tuple[str, str]:
    """Позволяет выбрать автомобиль по номеру из списка."""
    while True:
        value = ask_required("Выберите номер автомобиля из списка: ")
        if value.isdigit():
            index = int(value) - 1
            if 0 <= index < len(cars):
                car = cars.iloc[index]
                return str(car["car_name"]), str(car["car_number"])

        print("Введите корректный номер автомобиля из списка.")


def get_next_id(bookings: pd.DataFrame) -> int:
    """Определяет следующий id для новой брони."""
    if bookings.empty:
        return 1

    ids = pd.to_numeric(bookings["id"], errors="coerce").dropna()
    if ids.empty:
        return 1
    return int(ids.max()) + 1


def format_booking(booking: pd.Series) -> str:
    """Форматирует бронь для вывода в консоль и отчет."""
    return (
        f"ID {booking['id']}: {booking['car_name']} ({booking['car_number']}), "
        f"{booking['date']} {booking['start_time']}-{booking['end_time']}, "
        f"{booking['employee_name']}, {booking['route']}"
    )


def save_report(lines: list[str]) -> None:
    """Сохраняет результат операции в текстовый отчет."""
    OUTPUT_REPORT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_REPORT.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    configure_console_encoding()

    bookings = read_bookings()
    cars = show_cars(bookings)

    car_name, car_number = ask_car(cars)
    employee_name = ask_required("ФИО сотрудника: ")
    department = ask_required("Подразделение: ")
    date = ask_date()
    start_time, end_time = ask_time_interval()
    route = ask_required("Маршрут: ")
    purpose = ask_required("Цель поездки: ")

    conflicts = find_booking_conflicts(bookings, car_name, date, start_time, end_time)
    report_lines = [
        "Результат бронирования автомобиля",
        f"Автомобиль: {car_name} ({car_number})",
        f"Сотрудник: {employee_name}",
        f"Дата и время: {date} {start_time}-{end_time}",
        f"Маршрут: {route}",
        f"Цель: {purpose}",
        "",
    ]

    if conflicts.empty:
        new_booking = {
            "id": str(get_next_id(bookings)),
            "car_name": car_name,
            "car_number": car_number,
            "employee_name": employee_name,
            "department": department,
            "date": date,
            "start_time": start_time,
            "end_time": end_time,
            "route": route,
            "purpose": purpose,
            "status": "confirmed",
            "comment": "Создано через консольную программу",
        }

        bookings = pd.concat([bookings, pd.DataFrame([new_booking], columns=COLUMNS)], ignore_index=True)
        save_bookings(bookings)

        message = "Автомобиль свободен. Бронь добавлена со статусом confirmed."
        print(f"\n{message}")
        report_lines.append(message)
    else:
        message = "Автомобиль занят. Бронь не добавлена."
        print(f"\n{message}")
        print("\nКонфликтующие брони:")

        report_lines.append(message)
        report_lines.append("")
        report_lines.append("Конфликтующие брони:")

        for _, booking in conflicts.iterrows():
            formatted = format_booking(booking)
            print(formatted)
            report_lines.append(formatted)

    save_report(report_lines)
    print(f"\nОтчет сохранен: {OUTPUT_REPORT}")


if __name__ == "__main__":
    main()
