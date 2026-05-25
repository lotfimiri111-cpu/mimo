FROM python:3.12-slim

# Install system dependencies
RUN apt-get update && apt-get install -y \
    libreoffice \
    poppler-utils \
    fonts-liberation \
    fonts-dejavu \
    libglib2.0-0 \
    libsm6 \
    libxext6 \
    --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
RUN mkdir -p storage/pptx storage/receipts storage/preview_cache

ENV PORT=5000
EXPOSE 5000

RUN python3 -c "from core.database import init_db; init_db(); print('DB initialized')"

CMD ["bash", "start.sh"]
