from datetime import datetime

import pandas as pd


def build_datetime(date: str, time: str) -> datetime:
    """Собирает дату и время из строк Excel или HTML-формы."""
    return datetime.strptime(f"{date} {time}", "%Y-%m-%d %H:%M")


def find_booking_conflicts(
    bookings: pd.DataFrame,
    car_name: str,
    date: str,
    start_time: str,
    end_time: str,
    exclude_id: str | None = None,
) -> pd.DataFrame:
    """Возвращает подтвержденные брони, которые пересекаются с новым временем."""
    if bookings.empty:
        return bookings.copy()

    new_start = build_datetime(date, start_time)
    new_end = build_datetime(date, end_time)
    conflicts = []

    for _, booking in bookings.iterrows():
        # При редактировании или повторной проверке не считаем бронь конфликтом с самой собой.
        if exclude_id is not None and str(booking.get("id", "")) == str(exclude_id):
            continue

        # Блокируют время только подтвержденные брони того же автомобиля.
        if str(booking["car_name"]) != car_name:
            continue
        if str(booking["status"]) != "confirmed":
            continue

        existing_date = str(booking["date"])
        existing_start = build_datetime(existing_date, str(booking["start_time"]))
        existing_end = build_datetime(existing_date, str(booking["end_time"]))

        # Формула пересечения: new_start < existing_end AND new_end > existing_start.
        if new_start < existing_end and new_end > existing_start:
            conflicts.append(booking)

    return pd.DataFrame(conflicts)


def is_car_available(
    bookings: pd.DataFrame,
    car_name: str,
    date: str,
    start_time: str,
    end_time: str,
    exclude_id: str | None = None,
) -> bool:
    """Проверяет, свободен ли автомобиль в выбранный интервал."""
    return find_booking_conflicts(bookings, car_name, date, start_time, end_time, exclude_id).empty
