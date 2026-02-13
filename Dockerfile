# Dockerfile per a Flask en Cloud Run
FROM python:3.11-slim

# Evita que Python generi arxius .pyc i permet veure els logs en directe
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# Instal·lar dependències del sistema (si calguessin per a Pillow o docx, generalment no en slim per a aquestes)
# RUN apt-get update && apt-get install -y --no-install-recommends libmagic1 && rm -rf /var/lib/apt/lists/*

# Copiar requirements i instal·lar
COPY container/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt && pip install --no-cache-dir gunicorn

# Copiar el codi de l'aplicació
COPY container/ .

# El port que usa Cloud Run per defecte és el 8080
ENV PORT=8080

# Executar amb gunicorn
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 0 app:app
