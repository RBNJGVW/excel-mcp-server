FROM python:3.10-slim

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    FASTMCP_HOST=0.0.0.0 \
    FASTMCP_PORT=8017 \
    PYTHONPATH=/app/src

WORKDIR /app

# (opcional pero recomendado para certificados del sistema)
RUN apt-get update && apt-get install -y --no-install-recommends ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# Instala dependencias de Python desde requirements.txt (solo pip)
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copia el código (estructura con /src)
COPY src/ ./src/

EXPOSE 8017

# Arranca el servidor en SSE por defecto
CMD ["python", "-c", "from excel_mcp_server.server import run_sse; run_sse()"]