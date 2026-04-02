FROM node:22-alpine AS web-build

WORKDIR /app/ms-mail-fetcher-web

COPY ms-mail-fetcher-web/package.json ms-mail-fetcher-web/package-lock.json ./
RUN npm ci

COPY ms-mail-fetcher-web/ ./
RUN npm run build


FROM python:3.12-slim AS runtime

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    DATA_DIR=/data \
    HOST=0.0.0.0 \
    PORT=18765 \
    AUTO_PORT_FALLBACK=false

WORKDIR /app/ms-mail-fetcher-server

COPY ms-mail-fetcher-server/requirements.runtime.txt ./requirements.runtime.txt
RUN pip install --no-cache-dir -r requirements.runtime.txt

COPY ms-mail-fetcher-server/ ./
COPY --from=web-build /app/ms-mail-fetcher-web/dist ./template

EXPOSE 18765
VOLUME ["/data"]

CMD ["python", "app.py"]
