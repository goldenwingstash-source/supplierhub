# SupplierHub — онлайн-версия (без своего компьютера)

## Вариант 1. Render.com (бесплатно, 5 минут)

1. Зарегистрируйся на https://render.com (можно через GitHub или e-mail).
2. В панели: **New** → **Web Service** → **Deploy from a Git repo** не нужен:
   выбери **Public Git repo** или загрузи архив через **Upload** (если доступно).
   Если загрузки нет — создай пустой репозиторий на GitHub и залей туда файлы из этой папки.
3. Настройки:
   - Environment: **Python 3**
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `python server.py`
   - Plan: **Free**
4. Нажми **Create Web Service** → через 2–3 минуты получишь постоянный адрес вида
   `https://supplierhub-xxxx.onrender.com`.

## Вариант 2. Railway.app

1. https://railway.app → **New Project** → **Deploy from repo** (файлы этой папки).
2. Railway сам прочитает `Procfile` и `requirements.txt`.
3. Получишь адрес `https://xxxx.up.railway.app`.

## Вариант 3. Docker / свой сервер

```
docker build -t supplierhub .
docker run -d -p 8000:8000 supplierhub
```

## Важно

- Порт берётся из переменной окружения `PORT` (на Render/Railway задаётся автоматически).
- База `supplierhub.db` уже внутри — 5 544 позиции и 8 поставщиков.
- На бесплатном тарифе данные могут сбрасываться при перезапуске — для постоянного
  хранения нужно подключить диск (Render: Disk) или внешнюю БД.
- Телефон заходит по адресу из браузера, Wi-Fi и включённый ПК не нужны.
