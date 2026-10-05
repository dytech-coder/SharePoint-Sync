# SharePoint Sync: ein Image fuer beide Dienste (Web-App und Sync), siehe compose.yml
FROM python:3.14-slim

# tzdata: damit TZ=Europe/Berlin wirkt (Zeitstempel in Log und Lauf-Historie)
RUN apt-get update \
    && apt-get install -y --no-install-recommends tzdata \
    && rm -rf /var/lib/apt/lists/*

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# Programm nach /app - zuerst nur die Paketliste, damit Docker die installierten Pakete zwischenspeichert
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY sync.py web_app.py ./
COPY templates ./templates
COPY static ./static

# Daten (config.json, runs/, sharepoint_sync.log) liegen im Arbeitsverzeichnis /data = Volume.
# Nicht als root laufen; der Datenordner auf dem Host muss UID 1000 gehoeren.
RUN mkdir /data && chown 1000:1000 /data
USER 1000:1000
WORKDIR /data

EXPOSE 8000
CMD ["python", "/app/web_app.py"]
