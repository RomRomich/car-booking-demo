import os
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, flash, redirect, render_template, request, url_for

from booking_logic import find_booking_conflicts


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "local-car-booking-secret")

INPUT_FILE = Path("input/car_bookings.xlsx")

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


def read_bookings() -> pd.DataFrame:
    """Читает Excel и приводит все ячейки к строкам."""
    if not INPUT_FILE.exists():
        return pd.DataFrame(columns=COLUMNS)

    bookings = pd.read_excel(INPUT_FILE, dtype=str).fillna("")
    return bookings.reindex(columns=COLUMNS, fill_value="")


def save_bookings(bookings: pd.DataFrame) -> None:
    """Сохраняет брони обратно в Excel."""
    INPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    bookings.to_excel(INPUT_FILE, index=False)


def get_cars(bookings: pd.DataFrame) -> list[dict]:
    """Возвращает список автомобилей для фильтров и формы."""
    if bookings.empty:
        return []

    cars = bookings[["car_name", "car_number"]].drop_duplicates()
    cars = cars.sort_values(["car_name", "car_number"])
    return cars.to_dict("records")


def get_next_id(bookings: pd.DataFrame) -> str:
    """Определяет следующий id новой брони."""
    if bookings.empty:
        return "1"

    ids = pd.to_numeric(bookings["id"], errors="coerce").dropna()
    if ids.empty:
        return "1"
    return str(int(ids.max()) + 1)


def validate_interval(date: str, start_time: str, end_time: str) -> str | None:
    """Проверяет формат даты и времени."""
    try:
        start = datetime.strptime(f"{date} {start_time}", "%Y-%m-%d %H:%M")
        end = datetime.strptime(f"{date} {end_time}", "%Y-%m-%d %H:%M")
    except ValueError:
        return "Проверьте формат даты и времени."

    if start >= end:
        return "Время окончания должно быть позже времени начала."

    return None


def booking_to_dict(row: pd.Series) -> dict:
    """Преобразует строку DataFrame в словарь для шаблонов."""
    return {column: str(row.get(column, "")) for column in COLUMNS}


@app.route("/")
def index():
    """Главная страница с таблицей броней и фильтрами."""
    bookings = read_bookings()
    cars = get_cars(bookings)

    selected_date = request.args.get("date", "").strip()
    selected_car = request.args.get("car_name", "").strip()

    filtered = bookings.copy()
    if selected_date:
        filtered = filtered[filtered["date"] == selected_date]
    if selected_car:
        filtered = filtered[filtered["car_name"] == selected_car]

    filtered = filtered.sort_values(["date", "start_time", "car_name"])

    return render_template(
        "index.html",
        bookings=filtered.to_dict("records"),
        cars=cars,
        selected_date=selected_date,
        selected_car=selected_car,
    )


@app.route("/booking/new", methods=["GET", "POST"])
def new_booking():
    """Форма создания новой брони."""
    bookings = read_bookings()
    cars = get_cars(bookings)
    conflicts = []
    form_data = request.form.to_dict()

    if request.method == "POST":
        car_value = request.form.get("car", "")
        car_name, _, car_number = car_value.partition("|")
        employee_name = request.form.get("employee_name", "").strip()
        department = request.form.get("department", "").strip()
        date = request.form.get("date", "").strip()
        start_time = request.form.get("start_time", "").strip()
        end_time = request.form.get("end_time", "").strip()
        route = request.form.get("route", "").strip()
        purpose = request.form.get("purpose", "").strip()
        comment = request.form.get("comment", "").strip()

        required_values = [car_name, employee_name, department, date, start_time, end_time, route, purpose]
        if not all(required_values):
            flash("Заполните все обязательные поля.", "error")
        else:
            error = validate_interval(date, start_time, end_time)
            if error:
                flash(error, "error")
            else:
                conflicts_df = find_booking_conflicts(bookings, car_name, date, start_time, end_time)
                conflicts = conflicts_df.to_dict("records")

                if conflicts:
                    flash("Автомобиль занят. Бронь не добавлена.", "error")
                else:
                    new_row = {
                        "id": get_next_id(bookings),
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
                        "comment": comment,
                    }
                    bookings = pd.concat([bookings, pd.DataFrame([new_row], columns=COLUMNS)], ignore_index=True)
                    save_bookings(bookings)
                    flash("Бронь подтверждена", "success")
                    return redirect(url_for("index", date=date, car_name=car_name))

    return render_template("new_booking.html", cars=cars, conflicts=conflicts, form_data=form_data)


@app.post("/booking/<booking_id>/cancel")
def cancel_booking(booking_id: str):
    """Отменяет бронь через смену статуса на cancelled."""
    bookings = read_bookings()
    mask = bookings["id"].astype(str) == str(booking_id)

    if mask.any():
        bookings.loc[mask, "status"] = "cancelled"
        save_bookings(bookings)
        flash("Бронь отменена", "success")
    else:
        flash("Бронь не найдена", "error")

    return redirect(request.referrer or url_for("index"))


@app.route("/calendar")
def calendar():
    """Календарный вид: даты колонками, автомобили строками, брони блоками."""
    bookings = read_bookings()
    cars = get_cars(bookings)

    selected_date = request.args.get("date", "").strip()
    calendar_bookings = bookings.copy()
    if selected_date:
        calendar_bookings = calendar_bookings[calendar_bookings["date"] == selected_date]

    dates = sorted(calendar_bookings["date"].dropna().unique().tolist())
    grid = {}
    for car in cars:
        car_name = car["car_name"]
        grid[car_name] = {}
        for date in dates:
            day_bookings = calendar_bookings[
                (calendar_bookings["car_name"] == car_name) & (calendar_bookings["date"] == date)
            ]
            day_bookings = day_bookings.sort_values(["start_time", "end_time"])
            grid[car_name][date] = [booking_to_dict(row) for _, row in day_bookings.iterrows()]

    return render_template(
        "calendar.html",
        cars=cars,
        dates=dates,
        grid=grid,
        selected_date=selected_date,
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
