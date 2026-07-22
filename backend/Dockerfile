FROM python:3.12-slim

# Install LibreOffice (writer for docx, calc for xlsx, impress for pptx, headless)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-impress \
    libreoffice-core \
    fonts-dejavu \
    fonts-liberation \
    tesseract-ocr \
    tesseract-ocr-eng \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY main.py converter.py ./

# Where uploaded/converted files are temporarily stored
ENV CONVERTER_TMP_DIR=/tmp/converter_jobs
RUN mkdir -p /tmp/converter_jobs

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
