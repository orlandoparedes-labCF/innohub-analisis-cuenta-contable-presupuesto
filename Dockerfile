FROM python:3.11-slim

WORKDIR /app

# Instalar dependencias del sistema para Playwright
RUN apt-get update && apt-get install -y \
    wget gnupg curl \
    && rm -rf /var/lib/apt/lists/*

# Copiar y instalar dependencias Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Instalar browsers para Playwright (Clay enrichment)
RUN python -m playwright install chromium --with-deps

# Copiar código fuente
COPY . .

# Crear directorio de datos
RUN mkdir -p data

# Exponer puertos: 8000 = FastAPI (Hub), 8501 = Streamlit (UI)
EXPOSE 8000 8501

# Variables de entorno (valores reales via .env o variables del host)
ENV GROQ_API_KEY=""
ENV CLAY_EMAIL=""
ENV CLAY_PASSWORD=""

# Comando por defecto: iniciar API del Hub
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
