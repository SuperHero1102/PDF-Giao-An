# Ảnh Docker cho PDF Giáo Án: Python + poppler (đọc PDF) + Pandoc 3 (đổi công thức sang Word)
FROM python:3.12-slim
ARG PANDOC_VERSION=3.1.11.1
RUN apt-get update \
 && apt-get install -y --no-install-recommends poppler-utils curl ca-certificates fonts-liberation \
 && curl -fsSL -o /tmp/pandoc.deb "https://github.com/jgm/pandoc/releases/download/${PANDOC_VERSION}/pandoc-${PANDOC_VERSION}-1-$(dpkg --print-architecture).deb" \
 && apt-get install -y /tmp/pandoc.deb \
 && rm -rf /tmp/pandoc.deb /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV PORT=8000
EXPOSE 8000
# 1 worker + nhiều luồng: tiến trình nhận dạng nằm trong bộ nhớ của worker nên không chia nhiều worker
CMD gunicorn -w 1 --threads 8 -t 300 -b 0.0.0.0:${PORT} app:app
