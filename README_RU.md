# Система бронирования автомобилей

Локальное веб-приложение на Flask для бронирования автомобилей предприятия. Данные хранятся в Excel-файле `input/car_bookings.xlsx`.

## Структура проекта

```text
car_booking/
  input/
    car_bookings.xlsx
  output/
    booking_result.txt
  data/
  static/
    style.css
  templates/
    base.html
    index.html
    new_booking.html
    calendar.html
  app.py
  booking_logic.py
  main.py
  requirements.txt
  README_RU.md
```

## Подготовка на Windows

Откройте PowerShell в папке проекта:

```powershell
cd C:\CODEX\car_booking
```

Создайте виртуальное окружение:

```powershell
python -m venv .venv
```

Активируйте виртуальное окружение:

```powershell
.\.venv\Scripts\Activate.ps1
```

Установите зависимости:

```powershell
pip install -r requirements.txt
```

## Запуск веб-приложения

```powershell
python app.py
```

Приложение запускается на:

```text
http://127.0.0.1:5000
```

В коде используется:

```python
app.run(host="0.0.0.0", port=5000, debug=True)
```

Это позволяет открыть приложение не только на текущем компьютере, но и с других компьютеров в локальной сети.

## Открыть с другого компьютера в локальной сети

1. На компьютере, где запущено приложение, узнайте IP-адрес:

```powershell
ipconfig
```

2. Найдите адрес в строке `IPv4 Address`, например:

```text
192.168.1.25
```

3. На другом компьютере в той же сети откройте в браузере:

```text
http://192.168.1.25:5000
```

## Брандмауэр Windows

Чтобы открыть доступ с других компьютеров, нужно разрешить входящие подключения к порту `5000`.

Вариант через интерфейс Windows:

1. Откройте `Брандмауэр Защитника Windows в режиме повышенной безопасности`.
2. Перейдите в `Правила для входящих подключений`.
3. Создайте новое правило.
4. Выберите тип `Для порта`.
5. Укажите `TCP` и порт `5000`.
6. Разрешите подключение.
7. Примените правило для нужного профиля сети.
8. Назовите правило, например `Car Booking Flask 5000`.

Вариант через PowerShell от имени администратора:

```powershell
New-NetFirewallRule -DisplayName "Car Booking Flask 5000" -Direction Inbound -Protocol TCP -LocalPort 5000 -Action Allow
```

## Возможности веб-приложения

- таблица всех броней;
- фильтр по дате;
- фильтр по автомобилю;
- цветовые статусы: `confirmed`, `rejected`, `cancelled`;
- создание новой брони;
- проверка пересечений только по статусу `confirmed`;
- вывод конфликтующих броней;
- календарный вид по датам и автомобилям;
- отмена брони через смену статуса на `cancelled`.

## Старый консольный запуск

Консольная версия сохранена в файле `main.py`:

```powershell
python main.py
```

## Правило проверки пересечения

Бронь считается конфликтующей, если выполняется условие:

```text
new_start < existing_end AND new_end > existing_start
```

Статусы `cancelled` и `rejected` не блокируют автомобиль.
