FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY webapp ./webapp
COPY social_report ./social_report
RUN mkdir -p /app/runtime/fabric-home && chown -R 10001:10001 /app/runtime
ENV APP_DATA_DIR=/app/runtime HOME=/app/runtime/fabric-home PYTHONUNBUFFERED=1 PYTHONUTF8=1
USER 10001:10001
EXPOSE 8080
CMD ["python", "-m", "uvicorn", "webapp.app:app", "--host", "0.0.0.0", "--port", "8080"]
