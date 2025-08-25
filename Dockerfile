# Gunakan base image Python yang ringan
FROM python:3.10-slim

# Set direktori kerja di dalam container
WORKDIR /app

# Salin file requirements terlebih dahulu untuk caching
COPY requirements.txt .

# Install semua library yang dibutuhkan
RUN pip install --no-cache-dir -r requirements.txt

# Salin semua sisa file proyek
COPY . .

# Perintah yang akan dijalankan saat container启动
CMD ["python", "fullfitur.py"]
