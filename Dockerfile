FROM python:3.12-slim
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-writer libreoffice-calc libreoffice-impress fonts-noto-cjk && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY server.py index.html ./
RUN useradd --create-home --uid 10001 appuser && chown -R appuser:appuser /app
USER appuser
EXPOSE 10000
CMD ["uvicorn","server:app","--host","0.0.0.0","--port","10000"]
