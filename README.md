# Электронный дневник «Ценность» — запуск на Render

## 1. Загрузите проект на GitHub
1. Зарегистрируйтесь на github.com и создайте новый репозиторий (можно приватный).
2. Нажмите «uploading an existing file» и перетащите **содержимое** этой папки
   (app.py, requirements.txt, render.yaml, папки static и templates). Нажмите Commit.

## 2. Запустите на Render
1. Зайдите на render.com, войдите через GitHub.
2. New → **Blueprint** → выберите ваш репозиторий.
3. Render предложит ввести `TEACHER_PASSWORD` — придумайте пароль учителя (логин: `teacher`).
4. Нажмите Apply. Через несколько минут появится адрес вида `https://diary-xxxx.onrender.com`.

## 3. Вставьте на Тильду
Блок **T123 «HTML-код»**, вставьте:

    <iframe src="https://diary-xxxx.onrender.com" style="width:100%;height:1000px;border:0"></iframe>

Адрес замените на свой.

## Что важно знать
- Бесплатный веб-сервис Render «засыпает» без посещений, первая загрузка может занять около минуты.
- Бесплатная база PostgreSQL на Render ограничена по сроку и объёму — проверьте актуальные условия
  в своём аккаунте и делайте копии данных, если дневник рабочий.
- Safari и часть браузеров блокируют cookie внутри iframe: вход может не работать.
  Надёжное решение — подключить к Render поддомен вашего сайта (например, diary.вашсайт.ru:
  Render → Settings → Custom Domains) и использовать его в iframe.
- Локальный запуск: `pip install -r requirements.txt`, затем `python app.py` (используется SQLite).
