# Бронирование автомобилей

Flask-приложение для бронирования 3 автомобилей. Локально работает через SQLite, на Render может работать через PostgreSQL.

## Машины

В системе 3 автомобиля:

- HAVAL X585ВР550
- HAVAL Х588ВР550
- TOYOTA C011XK190

## Режим 1. Локальный тест

Установка:

```powershell
install.bat
```

Запуск:

```powershell
start.bat
```

Или вручную:

```powershell
python app.py
```

Адрес:

```text
http://127.0.0.1:5000
```

Локально используется:

```text
host=127.0.0.1
port=5000
database=data/car_booking.db
```

## Режим 2. Демо на Render

Render-версия используется для временного демо-теста. Чтобы данные сохранялись после перезапуска сервиса, обязательно подключите PostgreSQL и задайте `DATABASE_URL`.

Для Render:

```text
Build Command: pip install -r requirements.txt
Start Command: gunicorn app:app
```

Если Render передает переменную окружения `PORT`, приложение запускается на:

```text
host=0.0.0.0
port=$PORT
```

Если Render передает переменную окружения `DATABASE_URL`, приложение использует PostgreSQL через SQLAlchemy.

Если `DATABASE_URL` не задан, Render будет использовать SQLite-файл внутри web service. Такой вариант не подходит для теста с сохранением данных, потому что файловая система web service непостоянная.

Не загружайте реальные рабочие данные в Render-демо.

## Логотип

Логотип лежит здесь:

```text
static/img/georg_logo.png
```

Он используется в шапке сайта и как favicon.

## База данных

Локальная SQLite-база:

```text
data/car_booking.db
```

PostgreSQL на Render подключается через:

```text
DATABASE_URL
```

Для демо в базе лежат только тестовые записи без реальных сотрудников, маршрутов и конфиденциальной информации.

## Журнал действий

Страница:

```text
http://127.0.0.1:5000/logs
```

Журнал показывает создание, изменение, отмену, завершение поездки, импорт и экспорт Excel.

## Импорт Excel

Для реального локального теста старый Excel можно положить сюда:

```text
input/Резервирование автомобиля.xlsx
```

Этот файл добавлен в `.gitignore`, чтобы случайно не отправить реальные данные в GitHub и Render.

Для безопасного демо есть файл:

```text
input/demo_car_bookings.xlsx
```

## Экспорт Excel

Экспорт создается здесь:

```text
output/car_booking_report.xlsx
```

В Excel 3 листа, по одному на каждый автомобиль.

## Деплой

Подробная инструкция:

```text
README_DEPLOY_RU.md
```

## Резервная копия

Перед обновлением под логотип и Render создана папка:

```text
backup_before_logo_render_update
```
