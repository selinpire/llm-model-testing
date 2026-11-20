import os
import requests

def load_api_key():
    if os.path.exists(".env"):
        with open(".env", "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("OPENROUTER_API_KEY="):
                    return line.split("=", 1)[1].strip().strip('"').strip("'")
    return os.environ.get("OPENROUTER_API_KEY", "")

api_key = load_api_key()
if not api_key or "buraya_" in api_key:
    print("HATA: API Key .env dosyasında bulunamadı veya kaydedilmedi!")
    exit(1)

print(f"API Key yüklendi (başlangıç: {api_key[:10]}...)\n")

url = "https://openrouter.ai/api/v1/chat/completions"
headers = {
    "Authorization": f"Bearer {api_key}",
    "Content-Type": "application/json"
}

test_question = "Türkiye'nin başkenti neresidir? Sadece tek kelimeyle cevap ver."


# =========================================================
# 1. MODEL: deepseek/deepseek-v4.1-flash-1017
# =========================================================
model_1_name = "deepseek/deepseek-v4.1-flash-1017"
model_1_payload = {
    "model": model_1_name,
    "messages": [{"role": "user", "content": test_question}],
    "temperature": 0.6,
    "max_tokens": 500
}

try:
    r1 = requests.post(url, headers=headers, json=model_1_payload, timeout=30)
    if r1.status_code == 200:
        data1 = r1.json()
        ans1 = (data1["choices"][0]["message"].get("content") or "").strip()
        print(f"[1. MODEL BAŞARILI] {model_1_name} (temp={model_1_payload['temperature']}) -> {ans1}")
    else:
        print(f"[1. MODEL HATA {r1.status_code}] {model_1_name} -> {r1.text}")
except Exception as e:
    print(f"[1. MODEL İSTİSNA] {model_1_name} -> {e}")


# =========================================================
# 2. MODEL: mistralai/ministral-8b-2512
# =========================================================
model_2_name = "mistralai/ministral-8b-2512"
model_2_payload = {
    "model": model_2_name,
    "messages": [{"role": "user", "content": test_question}],
    "temperature": 0.5,
    "max_tokens": 500
}

try:
    r2 = requests.post(url, headers=headers, json=model_2_payload, timeout=30)
    if r2.status_code == 200:
        data2 = r2.json()
        ans2 = (data2["choices"][0]["message"].get("content") or "").strip()
        print(f"[2. MODEL BAŞARILI] {model_2_name} (temp={model_2_payload['temperature']}) -> {ans2}")
    else:
        print(f"[2. MODEL HATA {r2.status_code}] {model_2_name} -> {r2.text}")
except Exception as e:
    print(f"[2. MODEL İSTİSNA] {model_2_name} -> {e}")


# =========================================================
# 3. MODEL: google/gemma-3-12b-it
# =========================================================
model_3_name = "google/gemma-3-12b-it"
model_3_payload = {
    "model": model_3_name,
    "messages": [{"role": "user", "content": test_question}],
    "temperature": 0.7,
    "max_tokens": 500
}

try:
    r3 = requests.post(url, headers=headers, json=model_3_payload, timeout=30)
    if r3.status_code == 200:
        data3 = r3.json()
        ans3 = (data3["choices"][0]["message"].get("content") or "").strip()
        print(f"[3. MODEL BAŞARILI] {model_3_name} (temp={model_3_payload['temperature']}) -> {ans3}")
    else:
        print(f"[3. MODEL HATA {r3.status_code}] {model_3_name} -> {r3.text}")
except Exception as e:
    print(f"[3. MODEL İSTİSNA] {model_3_name} -> {e}")


# =========================================================
# HAKEM (JUDGE) MODEL: qwen/qwen3.8-flash
# =========================================================
judge_model_name = "qwen/qwen3.8-flash"
judge_payload = {
    "model": judge_model_name,
    "messages": [{"role": "user", "content": test_question}],
    "temperature": 0.0,  # Yorum yapmaması için 0.0
    "max_tokens": 500,
    "reasoning": {"effort": "none"}
}

try:
    r_judge = requests.post(url, headers=headers, json=judge_payload, timeout=30)
    if r_judge.status_code == 200:
        data_judge = r_judge.json()
        ans_judge = (data_judge["choices"][0]["message"].get("content") or "").strip()
        print(f"[HAKEM MODEL BAŞARILI] {judge_model_name} (temp={judge_payload['temperature']}) -> {ans_judge}")
    else:
        print(f"[HAKEM MODEL HATA {r_judge.status_code}] {judge_model_name} -> {r_judge.text}")
except Exception as e:
    print(f"[HAKEM MODEL İSTİSNA] {judge_model_name} -> {e}")
