FROM python:3.10-slim

RUN apt-get update && \
    apt-get install -y libreoffice curl && \
    pip install flask flask-cors gunicorn && \
    mkdir /app

WORKDIR /app
COPY . /app

CMD ["gunicorn", "main:app", "--bind", "0.0.0.0:5000"]
