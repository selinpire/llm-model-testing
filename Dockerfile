# Hafif ve resmi Python imajı
FROM python:3.11-slim

# Konteyner içi çalışma dizini
WORKDIR /app

# Python çıktılarının anlık olarak terminale yansıması için
ENV PYTHONUNBUFFERED=1

# Bağımlılıkları kopyala ve kur
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Proje dosyalarını kopyala
COPY create_dataset.py .
COPY test_connection.py .
COPY run_ab_test.py .
COPY soru_dataset.xlsx .

# Varsayılan olarak önce veri setini kontrol et/oluştur, ardından A/B testini çalıştır
CMD ["python", "run_ab_test.py", "--limit", "100"]
