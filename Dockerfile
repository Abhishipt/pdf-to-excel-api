FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Install system libraries required for camelot and Hindi fonts
RUN apt-get update && apt-get install -y \
    gcc \
    ghostscript \
    libglib2.0-0 \
    libsm6 \
    libxext6 \
    libxrender-dev \
    libgl1 \
    tesseract-ocr \
    poppler-utils \
    tcl \
    tk \
    fonts-noto-core \
    fonts-noto-cjk \
    fonts-noto-unhinted \
    fonts-noto-color-emoji \
 && apt-get clean \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY . /app

RUN pip install --upgrade pip
RUN pip install -r requirements.txt

EXPOSE 10000

CMD ["gunicorn", "--bind", "0.0.0.0:10000", "app:app"]
