# Dockerfile
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

# Dependencias nativas necesarias por tus libs (pyodbc, cryptography, etc.)
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
    build-essential unixodbc-dev libssl-dev libffi-dev ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# Instala deps Python
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

# Copia el código del proyecto (incluye run_sse.py y src/)
COPY . .

# Puerto donde expones FastMCP SSE
EXPOSE 8017

ENV FASTMCP_HOST=0.0.0.0 \
    FASTMCP_PORT=8017

CMD ["python", "-u", "run_sse.py"]