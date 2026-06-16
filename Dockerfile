FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY main.py .
# Render injects $PORT; bind to it. Single worker so the token cache + device
# flow are shared in one process (single-user home).
CMD uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000} --workers 1
