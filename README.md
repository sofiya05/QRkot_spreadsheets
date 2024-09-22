# QRKot

QRKot - это приложение для фонда, который собирает пожертвования на различные целевые проекты, связанные с поддержкой популяции кошек.

## Стек технологий

- **Python**
- **FastAPI**
- **FastAPI-Users**
- **SQLAlchemy**
- **Alembic**
- **Uvicorn**
- **Aiogoogle**

## Установка и запуск проекта

### Шаг 1: Клонируйте репозиторий

```bash
git clone https://github.com/your-repo/qrkot.git
cd qrkot
```

### Шаг 2: Создайте и активируйте виртуальное окружение

#### Для Windows:

```bash
python -m venv venv
venv\Scripts\activate
```

#### Для Unix / MacOS:

```bash
python3 -m venv venv
source venv/bin/activate
```

### Шаг 3: Обновите pip

```bash
python3 -m pip install --upgrade pip
```

### Шаг 4: Установите зависимости

```bash
pip install -r requirements.txt
```

### Шаг 5: Примените миграции базы данных

```bash
alembic upgrade head
```

### Шаг 6: Запустите приложение

```bash
uvicorn main:app --reload
```

### Документация проекта

После запуска проекта, документация будет доступна по адресам:

[http://127.0.0.1:8000/docs](http://127.0.0.1:8000/docs) - swagger
[http://127.0.0.1:8000/redoc](http://127.0.0.1:8000/redoc) - redoc

## Автор

Проект разработала **Бурсянина София**. Почта для связи [s.bursyanina@yandex.ru](mailto:s.bursyanina@yandex.ru)
