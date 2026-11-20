# 🧪 LLM-as-a-Judge Tabanlı A/B/n Model Karşılaştırma Testi

Bu proje, **OpenRouter API** üzerinden 3 farklı Büyük Dil Modelini (**LLM**) 100 soruluk Türkçe bir veri seti üzerinde yarıştıran ve modellerin verdiği yanıtları daha güçlü bir **Hakem Model (LLM-as-a-Judge)** aracılığıyla `0 - 100` arasında doğruluk olasılığına (%) göre değerlendiren kapsamlı bir A/B test sistemidir.

---

## 🎯 Projenin Amacı ve Mimarisi

1. **100 Soruluk Dengeli Veri Seti (`soru_dataset.xlsx`):**
   * **Kolay (1–35. Sorular):** Temel genel kültür, coğrafya, fen bilimleri ve dört işlem matematik.
   * **Orta (36–70. Sorular):** Tarih, edebiyat, biyoloji, fizik ve orta düzey matematik problemleri (denklem, yüzde, hız-zaman).
   * **Zor (71–100. Sorular):** Çok adımlı matematik işlemleri (kombinasyon, olasılık, bileşik faiz, işçi-havuz, ortalama hız), mantık problemleri ve detaylı genel kültür soruları.

2. **Test Edilen 3 Aday LLM Modeli:**
   | Sıra | Model Kısa Adı | OpenRouter Model ID | Temperature | Açıklama |
   | :---: | :--- | :--- | :---: | :--- |
   | **1** | **DeepSeek-v4.1-Flash** | `deepseek/deepseek-v4.1-flash-1017` | `0.6` | Yüksek akıl yürütme kapasiteli büyük model |
   | **2** | **Ministral-8B** | `mistralai/ministral-8b-2512` | `0.5` | Mistral ailesinin 8B parametreli kompakt modeli |
   | **3** | **Gemma-3-12B** | `google/gemma-3-12b-it` | `0.7` | Google ailesinin 12B parametreli açık ağırlıklı modeli |

3. **Karar Verici Hakem Model (Judge LLM):**
   * **Model:** `qwen/qwen3.8-flash`
   * **Temperature:** `0.0` (Yorum katmaması, tamamen deterministik ve nesnel puanlama yapması için `0.0` olarak sabitlenmiştir).
   * **Görev:** Her soru için `Soru`, `Ana Cevap (Referans)` ve `Modelin Cevabı` girdilerini karşılaştırarak **hiçbir yorum yapmadan yalnızca `%0` ile `%100` arasında bir doğruluk yüzdesi** üretir.

---

## 📊 Test Sonuçları Özeti (`genel_toplam_dogruluk.xlsx`)

100 soruluk test sonucunda elde edilen genel başarı tablosu:

| Model Sıra | Model Kısa Adı | OpenRouter Model ID | Temperature | Test Edilen Soru Sayısı | Tam Doğru (%100) Soru Sayısı | Genel Toplam Doğruluk Oranı (%) |
| :---: | :--- | :--- | :---: | :---: | :---: | :---: |
| **1** | **DeepSeek-v4.1-Flash** | `deepseek/deepseek-v4.1-flash-1017` | `0.6` | 100 | 100 | **%100.0** |
| **2** | **Ministral-8B** | `mistralai/ministral-8b-2512` | `0.5` | 100 | 77 | **%77.5** |
| **3** | **Gemma-3-12B** | `google/gemma-3-12b-it` | `0.7` | 100 | 80 | **%80.0** |

---

## 📂 Proje Dosya Yapısı

```text
llm_model_testing/
├── create_dataset.py                 # 100 soruluk Türkçe veri setini (soru_dataset.xlsx) oluşturur ve biçimlendirir
├── test_connection.py                # 3 aday model ve 1 hakem model için OpenRouter API bağlantı testi
├── run_ab_test.py                    # Ana A/B testi, hakem puanlaması ve Excel rapor üretim scripti
├── soru_dataset.xlsx                 # 100 soruluk (Soru No, Kategori, Soru, Ana Cevap) veri seti tablosu
├── model_cevaplari_ve_dogruluk.xlsx  # Her soru için 3 modelin cevabı ve hakem doğruluk yüzdeleri (%)
├── genel_toplam_dogruluk.xlsx        # 3 modelin genel toplam doğruluk oranlarını gösteren özet tablo
├── checkpoint_results.json           # Kesintilere karşı tüm soru ve cevapların JSON yedeği
├── requirements.txt                  # Gerekli Python kütüphaneleri (pandas, openpyxl, requests)
├── Dockerfile                        # Docker imaj yapılandırması
├── docker-compose.yml                # Docker Compose servis yapılandırması
├── .env.example                      # Örnek ortam değişkeni şablonu
└── .gitignore                        # Git tarafından yoksayılacak dosyalar (.env vb.)
```

---

## 🚀 Kurulum ve Çalıştırma

### 1. Ortam Değişkeninin (`.env`) Ayarlanması
Proje dizinindeki `.env.example` dosyasını `.env` olarak kopyalayın ve kendi OpenRouter API anahtarınızı girin:

```env
OPENROUTER_API_KEY=sk-or-v1-sizin_api_anahtariniz
```

### 2. Python ile Doğrudan Çalıştırma
```bash
# 1) Gerekli kütüphaneleri yükleyin
pip install -r requirements.txt

# 2) (İsteğe bağlı) Veri setini yeniden oluşturun
python create_dataset.py

# 3) API bağlantısını ve modelleri test edin
python test_connection.py

# 4) 100 soruluk A/B testini çalıştırın
python run_ab_test.py --limit 100
```

### 3. Docker ile Çalıştırma 🐳
Bilgisayarınızda Python kurulu olmasa bile Docker ile tek komutla çalıştırabilirsiniz:

```bash
# Docker Compose ile çalıştırma
docker compose up --build

# Veya klasik Docker komutlarıyla çalıştırma
docker build -t llm-ab-test .
docker run --rm --env-file .env -v "${PWD}:/app" llm-ab-test
```
