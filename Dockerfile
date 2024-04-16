# Используем базовый образ Python 3.9
FROM python:3.9

# Устанавливаем переменную окружения для Flask
ENV FLASK_APP=app.py

# Копируем все файлы из текущего каталога в каталог /app в контейнере
COPY . /app

# Копируем requirements.txt и устанавливаем зависимости Python
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

# Определяем рабочую директорию
WORKDIR /app

# Команда для запуска Flask приложения
CMD ["flask", "run", "--host=0.0.0.0"]