# Dockerfile para despliegue en Google Cloud Run
FROM python:3.11-slim

# Evitar que Python escriba archivos .pyc y activar salida sin búfer
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# Instalar dependencias del sistema necesarias para compilar o procesar documentos si fuera necesario
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    libpq-dev \
    && rm -rf /var/lib/apt/lists/*

# Copiar requirements e instalar dependencias de Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar todo el código de la aplicación
COPY . .

# Cloud Run inyecta la variable de entorno PORT (por defecto 8080)
ENV PORT=8080

# Ejecutar Gunicorn configurado para Cloud Run
# timeout 0 recomendado para Cloud Run para permitir peticiones largas (generación de documentos)
CMD exec gunicorn --bind 0.0.0.0:$PORT --workers 1 --threads 8 --timeout 0 app:app
