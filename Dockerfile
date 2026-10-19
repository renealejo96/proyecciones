FROM python:3.11-slim

WORKDIR /app

# Install dependencies
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY . .

# Expose internal Flask port
EXPOSE 5000

ENV FLASK_APP=app.py
ENV FLASK_DEBUG=0
ENV PYTHONUNBUFFERED=1

# Run migrations once in an isolated process, then start Gunicorn
CMD ["sh", "-c", "python -c 'from db import init_db; init_db()' && gunicorn --bind 0.0.0.0:5000 --workers 4 --timeout 120 app:app"]
