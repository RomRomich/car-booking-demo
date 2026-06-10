# Повторный деплой демо на Render с PostgreSQL

Это демо-версия для теста на 1-4 недели. Не используйте ее для реальных рабочих данных.

## Проверить локально

```powershell
python app.py
```

Откройте:

```text
http://127.0.0.1:5000
```

Локально, если `DATABASE_URL` не задан, используется SQLite:

```text
data/car_booking.db
```

## Залить изменения на GitHub

```powershell
git status
git add .
git commit -m "Update car booking demo with PostgreSQL"
git push origin main
```

## Создать PostgreSQL Database в Render

1. Откройте Render Dashboard.
2. Нажмите `New +`.
3. Выберите `PostgreSQL`.
4. Создайте базу данных для демо.
5. После создания откройте страницу базы.
6. Скопируйте `Internal Database URL`.

## Подключить PostgreSQL к Web Service

1. Откройте Render Dashboard.
2. Выберите сервис `car-booking-demo`.
3. Откройте `Environment`.
4. Добавьте переменную:

```text
DATABASE_URL = Internal Database URL
```

Render может дать адрес вида `postgres://...`. Приложение автоматически заменяет его на `postgresql://...` для SQLAlchemy.

## Настройки Web Service

Репозиторий:

```text
RomRomich/car-booking-demo
```

Ветка:

```text
main
```

Build Command:

```text
pip install -r requirements.txt
```

Start Command:

```text
gunicorn app:app
```

После настройки нажмите:

```text
Manual Deploy -> Deploy latest commit
```

Или дождитесь автодеплоя после `git push`.

## Что проверить после деплоя

- открывается главная страница;
- виден крупный логотип Георг Полимер;
- создается новая бронь;
- изменения брони сохраняются;
- отмена сохраняется;
- завершение поездки сохраняется;
- журнал действий работает;
- календарь работает;
- экспорт Excel работает;
- даты отображаются в формате `ДД.ММ.ГГГГ`;
- после перезапуска сервиса данные остаются в PostgreSQL.

## Предупреждение

Render-версия предназначена только для демонстрационного теста. Для реальной работы внутри Георг Полимер используйте локальный запуск в сети предприятия или внутренний сервер.
