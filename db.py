import os
from datetime import datetime
from pathlib import Path

from sqlalchemy import (
    Boolean,
    Column,
    DateTime,
    ForeignKey,
    Integer,
    MetaData,
    String,
    Table,
    Text,
    and_,
    create_engine,
    inspect,
    or_,
    select,
    text,
)


DB_PATH = Path("data/car_booking.db")

CARS = [
    ("HAVAL X585ВР550", "X585ВР550"),
    ("HAVAL Х588ВР550", "Х588ВР550"),
    ("TOYOTA C011XK190", "C011XK190"),
]

BLOCKING_STATUSES = ("confirmed", "active")
STATUSES = ("confirmed", "active", "completed", "cancelled")
LOG_ACTIONS = ("created", "updated", "cancelled", "completed", "imported", "exported")


def get_database_url() -> str:
    """Возвращает URL базы: PostgreSQL на Render или SQLite локально."""
    database_url = os.environ.get("DATABASE_URL", "").strip()
    if database_url.startswith("postgres://"):
        database_url = database_url.replace("postgres://", "postgresql://", 1)
    if database_url:
        return database_url

    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    return f"sqlite:///{DB_PATH.as_posix()}"


DATABASE_URL = get_database_url()
engine = create_engine(DATABASE_URL, future=True, pool_pre_ping=True)
metadata = MetaData()

cars_table = Table(
    "cars",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("name", String(255), nullable=False),
    Column("number", String(255), nullable=False),
    Column("is_active", Boolean, nullable=False, default=True),
)

bookings_table = Table(
    "bookings",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("car_id", Integer, ForeignKey("cars.id"), nullable=False),
    Column("employee_name", String(255), nullable=False),
    Column("start_date", String(10), nullable=False),
    Column("start_time", String(5), nullable=True),
    Column("end_date", String(10), nullable=False),
    Column("end_time", String(5), nullable=True),
    Column("route", Text, nullable=True),
    Column("odometer_return", String(255), nullable=True),
    Column("status", String(50), nullable=False),
    Column("comment", Text, nullable=True),
    Column("created_by", String(255), nullable=True),
    Column("updated_by", String(255), nullable=True),
    Column("created_at", DateTime, nullable=False),
    Column("updated_at", DateTime, nullable=False),
    Column("cancelled_at", DateTime, nullable=True),
    Column("purpose", Text, nullable=True),
    Column("department", Text, nullable=True),
)

activity_log_table = Table(
    "activity_log",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("action", String(50), nullable=False),
    Column("booking_id", Integer, nullable=True),
    Column("car_name", String(255), nullable=True),
    Column("actor_name", String(255), nullable=True),
    Column("employee_name", String(255), nullable=True),
    Column("details", Text, nullable=True),
    Column("created_at", DateTime, nullable=False),
)


def now_dt() -> datetime:
    """Возвращает текущие дату и время."""
    return datetime.now()


def row_to_dict(row) -> dict:
    """Преобразует SQLAlchemy Row в обычный словарь."""
    return dict(row._mapping)


def normalize_datetime(value) -> str:
    """Возвращает дату-время в строковом формате для совместимости с шаблонами."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    return str(value)


def normalize_booking_row(row: dict) -> dict:
    """Нормализует типы строки брони для существующего app.py."""
    for field in ("created_at", "updated_at", "cancelled_at"):
        row[field] = normalize_datetime(row.get(field))
    for field in ("start_time", "end_time", "route", "odometer_return", "comment", "purpose", "department"):
        row[field] = row.get(field) or ""
    return row


def normalize_log_row(row: dict) -> dict:
    """Нормализует типы строки журнала."""
    row["created_at"] = normalize_datetime(row.get("created_at"))
    for field in ("car_name", "actor_name", "employee_name", "details"):
        row[field] = row.get(field) or ""
    return row


def get_connection():
    """Возвращает SQLAlchemy connection для служебных проверок."""
    return engine.connect()


def table_columns(table_name: str) -> set[str]:
    """Возвращает список колонок существующей таблицы."""
    inspector = inspect(engine)
    if not inspector.has_table(table_name):
        return set()
    return {column["name"] for column in inspector.get_columns(table_name)}


def sql_type_for_column(column_name: str) -> str:
    """Возвращает простой SQL-тип для миграции SQLite/PostgreSQL."""
    if column_name in {"created_at", "updated_at", "cancelled_at"}:
        return "TIMESTAMP"
    return "TEXT"


def add_missing_columns() -> None:
    """Добавляет недостающие колонки в старую локальную SQLite-базу или старую демо-базу."""
    expected = {
        "bookings": {
            "start_time",
            "end_time",
            "purpose",
            "department",
            "created_by",
            "updated_by",
        },
    }
    with engine.begin() as connection:
        for table_name, columns in expected.items():
            existing = table_columns(table_name)
            for column_name in columns - existing:
                connection.execute(text(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {sql_type_for_column(column_name)}"))


def init_db() -> None:
    """Создает таблицы и справочник из трех автомобилей."""
    metadata.create_all(engine)
    add_missing_columns()

    with engine.begin() as connection:
        for index, (name, number) in enumerate(CARS, start=1):
            existing = connection.execute(select(cars_table.c.id).where(cars_table.c.id == index)).first()
            values = {"id": index, "name": name, "number": number, "is_active": True}
            if existing is None:
                connection.execute(cars_table.insert().values(**values))
            else:
                connection.execute(
                    cars_table.update()
                    .where(cars_table.c.id == index)
                    .values(name=name, number=number, is_active=True)
                )


def fetch_cars(active_only: bool = True) -> list[dict]:
    """Возвращает список автомобилей."""
    init_db()
    query = select(cars_table).order_by(cars_table.c.id)
    if active_only:
        query = query.where(cars_table.c.is_active.is_(True))
    with engine.connect() as connection:
        return [row_to_dict(row) for row in connection.execute(query)]


def count_cars() -> int:
    """Возвращает количество активных автомобилей."""
    return len(fetch_cars())


def booking_select():
    """Базовый SELECT брони с данными автомобиля."""
    return select(
        bookings_table,
        cars_table.c.name.label("car_name"),
        cars_table.c.number.label("car_number"),
    ).select_from(bookings_table.join(cars_table, cars_table.c.id == bookings_table.c.car_id))


def fetch_booking(booking_id: int) -> dict | None:
    """Возвращает одну бронь."""
    init_db()
    query = booking_select().where(bookings_table.c.id == booking_id)
    with engine.connect() as connection:
        row = connection.execute(query).first()
    return normalize_booking_row(row_to_dict(row)) if row else None


def fetch_bookings(filters: dict | None = None) -> list[dict]:
    """Возвращает брони с фильтрами."""
    init_db()
    filters = filters or {}
    conditions = []

    if filters.get("car_id"):
        conditions.append(bookings_table.c.car_id == int(filters["car_id"]))
    if filters.get("status"):
        conditions.append(bookings_table.c.status == filters["status"])
    if filters.get("date_from"):
        conditions.append(bookings_table.c.end_date >= filters["date_from"])
    if filters.get("date_to"):
        conditions.append(bookings_table.c.start_date <= filters["date_to"])
    if filters.get("search"):
        search = f"%{filters['search']}%"
        conditions.append(or_(bookings_table.c.employee_name.like(search), bookings_table.c.route.like(search)))

    query = booking_select()
    if conditions:
        query = query.where(and_(*conditions))
    query = query.order_by(bookings_table.c.start_date.desc(), bookings_table.c.start_time.desc(), bookings_table.c.id.desc())

    with engine.connect() as connection:
        return [normalize_booking_row(row_to_dict(row)) for row in connection.execute(query)]


def is_multiday(start_date: str, end_date: str) -> bool:
    """Проверяет, занимает ли бронь несколько дней."""
    return start_date != end_date


def times_are_empty(start_time: str | None, end_time: str | None) -> bool:
    """Проверяет, что оба времени не заполнены."""
    return not (start_time or "").strip() and not (end_time or "").strip()


def times_overlap(new_start: str, new_end: str, existing_start: str, existing_end: str) -> bool:
    """Формула пересечения времени."""
    return new_start < existing_end and new_end > existing_start


def booking_conflicts_with_candidate(candidate: dict, existing: dict) -> bool:
    """Проверяет конфликт по датам и времени."""
    if candidate["end_date"] < existing["start_date"] or candidate["start_date"] > existing["end_date"]:
        return False
    if is_multiday(candidate["start_date"], candidate["end_date"]):
        return True
    if is_multiday(existing["start_date"], existing["end_date"]):
        return True
    if times_are_empty(candidate.get("start_time"), candidate.get("end_time")):
        return True
    if times_are_empty(existing.get("start_time"), existing.get("end_time")):
        return True
    return times_overlap(candidate["start_time"], candidate["end_time"], existing["start_time"], existing["end_time"])


def find_conflicts(
    car_id: int,
    start_date: str,
    end_date: str,
    start_time: str = "",
    end_time: str = "",
    exclude_booking_id: int | None = None,
) -> list[dict]:
    """Ищет конфликтующие confirmed/active брони."""
    filters = {"car_id": str(car_id), "date_from": start_date, "date_to": end_date}
    candidates = [booking for booking in fetch_bookings(filters) if booking["status"] in BLOCKING_STATUSES]
    if exclude_booking_id is not None:
        candidates = [booking for booking in candidates if int(booking["id"]) != int(exclude_booking_id)]

    candidate = {
        "start_date": start_date,
        "end_date": end_date,
        "start_time": start_time,
        "end_time": end_time,
    }
    return [booking for booking in candidates if booking_conflicts_with_candidate(candidate, booking)]


def create_booking(data: dict) -> int:
    """Создает новую бронь."""
    timestamp = now_dt()
    actor_name = data.get("actor_name", "") or data.get("employee_name", "")
    values = {
        "car_id": data["car_id"],
        "employee_name": data["employee_name"],
        "start_date": data["start_date"],
        "start_time": data.get("start_time", ""),
        "end_date": data["end_date"],
        "end_time": data.get("end_time", ""),
        "route": data.get("route", ""),
        "odometer_return": data.get("odometer_return", ""),
        "status": data.get("status", "confirmed"),
        "comment": data.get("comment", ""),
        "created_by": actor_name,
        "updated_by": actor_name,
        "created_at": timestamp,
        "updated_at": timestamp,
        "cancelled_at": None,
        "purpose": "",
        "department": "",
    }
    with engine.begin() as connection:
        result = connection.execute(bookings_table.insert().values(**values))
        return int(result.inserted_primary_key[0])


def update_booking(booking_id: int, data: dict) -> None:
    """Обновляет бронь."""
    actor_name = data.get("actor_name", "") or data.get("employee_name", "")
    values = {
        "car_id": data["car_id"],
        "employee_name": data["employee_name"],
        "start_date": data["start_date"],
        "start_time": data.get("start_time", ""),
        "end_date": data["end_date"],
        "end_time": data.get("end_time", ""),
        "route": data.get("route", ""),
        "odometer_return": data.get("odometer_return", ""),
        "status": data["status"],
        "comment": data.get("comment", ""),
        "updated_by": actor_name,
        "updated_at": now_dt(),
        "purpose": "",
        "department": "",
    }
    with engine.begin() as connection:
        connection.execute(bookings_table.update().where(bookings_table.c.id == booking_id).values(**values))


def set_booking_status(
    booking_id: int,
    status: str,
    odometer_return: str | None = None,
    extra_comment: str | None = None,
) -> None:
    """Меняет статус брони."""
    booking = fetch_booking(booking_id)
    old_comment = booking["comment"] if booking else ""
    new_comment = f"{old_comment}\n{extra_comment}".strip() if extra_comment else old_comment
    values = {
        "status": status,
        "comment": new_comment,
        "updated_at": now_dt(),
    }
    if odometer_return is not None:
        values["odometer_return"] = odometer_return
    if status == "cancelled":
        values["cancelled_at"] = now_dt()

    with engine.begin() as connection:
        connection.execute(bookings_table.update().where(bookings_table.c.id == booking_id).values(**values))


def log_activity(
    action: str,
    booking_id: int | None = None,
    car_name: str = "",
    actor_name: str = "",
    employee_name: str = "",
    details: str = "",
) -> None:
    """Добавляет запись в журнал действий."""
    init_db()
    values = {
        "action": action,
        "booking_id": booking_id,
        "car_name": car_name,
        "actor_name": actor_name,
        "employee_name": employee_name,
        "details": details,
        "created_at": now_dt(),
    }
    with engine.begin() as connection:
        connection.execute(activity_log_table.insert().values(**values))


def fetch_logs(filters: dict | None = None) -> list[dict]:
    """Возвращает журнал действий с фильтрами."""
    init_db()
    filters = filters or {}
    conditions = []

    if filters.get("action"):
        conditions.append(activity_log_table.c.action == filters["action"])
    if filters.get("car_name"):
        conditions.append(activity_log_table.c.car_name == filters["car_name"])
    if filters.get("date_from"):
        conditions.append(activity_log_table.c.created_at >= datetime.strptime(filters["date_from"], "%Y-%m-%d"))
    if filters.get("date_to"):
        end = datetime.strptime(filters["date_to"], "%Y-%m-%d").replace(hour=23, minute=59, second=59)
        conditions.append(activity_log_table.c.created_at <= end)
    if filters.get("search"):
        search = f"%{filters['search']}%"
        conditions.append(or_(activity_log_table.c.employee_name.like(search), activity_log_table.c.details.like(search)))

    query = select(activity_log_table)
    if conditions:
        query = query.where(and_(*conditions))
    query = query.order_by(activity_log_table.c.created_at.desc(), activity_log_table.c.id.desc())

    with engine.connect() as connection:
        return [normalize_log_row(row_to_dict(row)) for row in connection.execute(query)]
