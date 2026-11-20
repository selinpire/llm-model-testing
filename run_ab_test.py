import os
import re
import json
import time
import argparse
import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# =========================================================
# 1. MODEL AYARLARI: deepseek/deepseek-v4.1-flash-1017
# =========================================================
MODEL_1_SHORT_NAME = "DeepSeek-v4.1-Flash"
MODEL_1_ID = "deepseek/deepseek-v4.1-flash-1017"
MODEL_1_TEMPERATURE = 0.6
MODEL_1_MAX_TOKENS = 600

# =========================================================
# 2. MODEL AYARLARI: mistralai/ministral-8b-2512
# =========================================================
MODEL_2_SHORT_NAME = "Ministral-8B"
MODEL_2_ID = "mistralai/ministral-8b-2512"
MODEL_2_TEMPERATURE = 0.5
MODEL_2_MAX_TOKENS = 600

# =========================================================
# 3. MODEL AYARLARI: google/gemma-3-12b-it
# =========================================================
MODEL_3_SHORT_NAME = "Gemma-3-12B"
MODEL_3_ID = "google/gemma-3-12b-it"
MODEL_3_TEMPERATURE = 0.7
MODEL_3_MAX_TOKENS = 600

# =========================================================
# HAKEM (JUDGE) MODEL AYARLARI: qwen/qwen3.8-flash
# =========================================================
JUDGE_MODEL_ID = "qwen/qwen3.8-flash"
JUDGE_TEMPERATURE = 0.0  # Yorum yapmaması için 0.0
JUDGE_MAX_TOKENS = 100


OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"

def load_api_key():
    if os.path.exists(".env"):
        with open(".env", "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("OPENROUTER_API_KEY="):
                    return line.split("=", 1)[1].strip().strip('"').strip("'")
    return os.environ.get("OPENROUTER_API_KEY", "")

API_KEY = load_api_key()
HEADERS = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type": "application/json"
}

SYSTEM_PROMPT_CANDIDATE = (
    "Sen bilgili bir asistansın. Sorulan sorulara Türkçe, doğrudan, kısa ve net cevap ver. "
    "Gereksiz uzun açıklamalardan kaçın."
)


# ---------------------------------------------------------
# 1. MODEL FONKSİYONU (deepseek/deepseek-v4.1-flash-1017)
# ---------------------------------------------------------
def ask_model_1(question):
    payload = {
        "model": MODEL_1_ID,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT_CANDIDATE},
            {"role": "user", "content": question}
        ],
        "temperature": MODEL_1_TEMPERATURE,
        "max_tokens": MODEL_1_MAX_TOKENS
    }
    for attempt in range(1, 7):
        try:
            resp = requests.post(OPENROUTER_URL, headers=HEADERS, json=payload, timeout=45)
            if resp.status_code == 200:
                content = (resp.json()["choices"][0]["message"].get("content") or "").strip()
                if content:
                    return content
                payload["max_tokens"] = 1200
            elif resp.status_code == 429:
                time.sleep(3 * attempt)
            else:
                time.sleep(2 * attempt)
        except Exception as e:
            if attempt == 6:
                return f"HATA: {e}"
            time.sleep(2 * attempt)
    return "HATA: Yanıt alınamadı"


# ---------------------------------------------------------
# 2. MODEL FONKSİYONU (mistralai/ministral-8b-2512)
# ---------------------------------------------------------
def ask_model_2(question):
    payload = {
        "model": MODEL_2_ID,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT_CANDIDATE},
            {"role": "user", "content": question}
        ],
        "temperature": MODEL_2_TEMPERATURE,
        "max_tokens": MODEL_2_MAX_TOKENS
    }
    for attempt in range(1, 7):
        try:
            resp = requests.post(OPENROUTER_URL, headers=HEADERS, json=payload, timeout=45)
            if resp.status_code == 200:
                content = (resp.json()["choices"][0]["message"].get("content") or "").strip()
                if content:
                    return content
                payload["max_tokens"] = 1200
            elif resp.status_code == 429:
                time.sleep(3 * attempt)
            else:
                time.sleep(2 * attempt)
        except Exception as e:
            if attempt == 6:
                return f"HATA: {e}"
            time.sleep(2 * attempt)
    return "HATA: Yanıt alınamadı"


# ---------------------------------------------------------
# 3. MODEL FONKSİYONU (google/gemma-3-12b-it)
# ---------------------------------------------------------
def ask_model_3(question):
    payload = {
        "model": MODEL_3_ID,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT_CANDIDATE},
            {"role": "user", "content": question}
        ],
        "temperature": MODEL_3_TEMPERATURE,
        "max_tokens": MODEL_3_MAX_TOKENS
    }
    for attempt in range(1, 7):
        try:
            resp = requests.post(OPENROUTER_URL, headers=HEADERS, json=payload, timeout=45)
            if resp.status_code == 200:
                content = (resp.json()["choices"][0]["message"].get("content") or "").strip()
                if content:
                    return content
                payload["max_tokens"] = 1200
            elif resp.status_code == 429:
                time.sleep(3 * attempt)
            else:
                time.sleep(2 * attempt)
        except Exception as e:
            if attempt == 6:
                return f"HATA: {e}"
            time.sleep(2 * attempt)
    return "HATA: Yanıt alınamadı"


# ---------------------------------------------------------
# HAKEM (JUDGE) MODEL FONKSİYONU (qwen/qwen3.8-flash)
# ---------------------------------------------------------
def evaluate_with_judge(question, reference_answer, model_answer):
    system_prompt = (
        "Sen nesnel ve adil bir cevap değerlendirme hakemisin (LLM Judge). "
        "Sana bir Soru, bu sorunun doğru Ana Cevabı ve bir yapay zeka modelinin verdiği Model Cevabı verilecek. "
        "Modelin verdiği cevabın Ana Cevaba göre ne kadar doğru olduğunu yüzde (%) olarak 0 ile 100 arasında bir tam sayı ile puanla.\n"
        "Kurallar:\n"
        "1. Eğer modelin cevabı ana cevabı doğru şekilde içeriyorsa (cümle içinde veya ekstra doğru açıklamayla vermiş olsa bile) 100 ver.\n"
        "2. Eğer soru birden fazla bilgi istiyorsa ve model sadece bir kısmını doğru bildiyse, veya matematiksel işlem yöntemini doğru kurup son adımda küçük bir hesap/yuvarlama hatası yaptıysa doğruluğuna göre 25, 50, 75 gibi kısmi bir puan ver.\n"
        "3. Tamamen yanlışsa veya alakasızsa 0 ver.\n"
        "4. KESİNLİKLE hiçbir yorum, açıklama veya kelime yazma! Sadece 0 ile 100 arasında tek bir sayı döndür (örneğin: 100 veya 50 veya 0)."
    )
    user_prompt = (
        f"Soru: {question}\n"
        f"Ana Cevap: {reference_answer}\n"
        f"Model Cevabı: {model_answer}\n\n"
        f"Doğruluk Oranı (0-100 arası sadece sayı):"
    )
    payload = {
        "model": JUDGE_MODEL_ID,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": JUDGE_TEMPERATURE,
        "max_tokens": JUDGE_MAX_TOKENS,
        "reasoning": {"effort": "none"}
    }

    raw_score = ""
    for attempt in range(1, 5):
        try:
            resp = requests.post(OPENROUTER_URL, headers=HEADERS, json=payload, timeout=45)
            if resp.status_code == 200:
                raw_score = (resp.json()["choices"][0]["message"].get("content") or "").strip()
                if raw_score:
                    break
            time.sleep(attempt)
        except Exception:
            time.sleep(attempt)

    matches = re.findall(r"\b(100|[1-9]?\d)\b", raw_score)
    if matches:
        return int(matches[0])
    digits = re.findall(r"\d+", raw_score)
    if digits:
        return max(0, min(100, int(digits[0])))
    return 0


def process_single_question(row_dict):
    q_no = int(row_dict["Soru No"])
    cat = row_dict["Kategori"]
    question = row_dict["Soru"]
    ref_ans = str(row_dict["Ana Cevap"])

    # 1. Model Cevabı ve Hakem Puanı
    ans_1 = ask_model_1(question)
    score_1 = evaluate_with_judge(question, ref_ans, ans_1)

    # 2. Model Cevabı ve Hakem Puanı
    ans_2 = ask_model_2(question)
    score_2 = evaluate_with_judge(question, ref_ans, ans_2)

    # 3. Model Cevabı ve Hakem Puanı
    ans_3 = ask_model_3(question)
    score_3 = evaluate_with_judge(question, ref_ans, ans_3)

    return {
        "Soru No": q_no,
        "Kategori": cat,
        "Soru": question,
        "Ana Cevap": ref_ans,
        f"{MODEL_1_SHORT_NAME} Cevap": ans_1,
        f"{MODEL_1_SHORT_NAME} Doğruluk (%)": score_1,
        f"{MODEL_2_SHORT_NAME} Cevap": ans_2,
        f"{MODEL_2_SHORT_NAME} Doğruluk (%)": score_2,
        f"{MODEL_3_SHORT_NAME} Cevap": ans_3,
        f"{MODEL_3_SHORT_NAME} Doğruluk (%)": score_3,
    }


def format_detailed_excel(filepath):
    wb = load_workbook(filepath)
    ws = wb.active
    ws.title = "Model Cevapları ve Doğruluk"

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    zebra_fill = PatternFill(start_color="F2F6FA", end_color="F2F6FA", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    red_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    max_col = ws.max_column
    max_row = ws.max_row

    ws.row_dimensions[1].height = 30
    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for row_idx in range(2, max_row + 1):
        ws.row_dimensions[row_idx].height = 28
        row_fill = zebra_fill if row_idx % 2 == 0 else white_fill
        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            col_name = ws.cell(row=1, column=col_idx).value or ""

            if "Doğruluk (%)" in col_name:
                val = cell.value if isinstance(cell.value, (int, float)) else 0
                cell.value = f"%{int(val)}"
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(name="Calibri", size=11, bold=True, color="000000")
                if val >= 80:
                    cell.fill = green_fill
                elif val >= 40:
                    cell.fill = yellow_fill
                else:
                    cell.fill = red_fill
            elif col_idx == 1:
                cell.fill = row_fill
                cell.font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 2:
                cell.fill = row_fill
                cell.font = Font(name="Calibri", size=11)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 4:
                cell.fill = row_fill
                cell.font = Font(name="Calibri", size=11, bold=True)
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                cell.fill = row_fill
                cell.font = Font(name="Calibri", size=11)
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    widths = {
        "A": 10, "B": 22, "C": 55, "D": 25,
        "E": 35, "F": 22, "G": 35, "H": 22, "I": 35, "J": 22,
    }
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w

    ws.freeze_panes = "A2"
    wb.save(filepath)


def format_summary_excel(filepath):
    wb = load_workbook(filepath)
    ws = wb.active
    ws.title = "Genel Toplam Doğruluk"

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    zebra_fill = PatternFill(start_color="F2F6FA", end_color="F2F6FA", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    max_col = ws.max_column
    max_row = ws.max_row

    ws.row_dimensions[1].height = 28
    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for row_idx in range(2, max_row + 1):
        ws.row_dimensions[row_idx].height = 26
        row_fill = zebra_fill if row_idx % 2 == 0 else white_fill
        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill = row_fill
            cell.border = thin_border
            if col_idx in (1, 2):
                cell.font = Font(name="Calibri", size=11, bold=(col_idx == 2))
                cell.alignment = Alignment(horizontal="left" if col_idx > 1 else "center", vertical="center")
            else:
                cell.font = Font(name="Calibri", size=11, bold=(col_idx == max_col))
                cell.alignment = Alignment(horizontal="center", vertical="center")

    widths = {"A": 12, "B": 25, "C": 42, "D": 16, "E": 18, "F": 22, "G": 28}
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w

    wb.save(filepath)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--start", type=int, default=1, help="Başlangıç soru numarası (1-100)")
    parser.add_argument("--limit", type=int, default=100, help="Kaç soru test edilecek")
    parser.add_argument("--prefix", type=str, default="", help="Dosya adı ön eki")
    args = parser.parse_args()

    df_questions = pd.read_excel("soru_dataset.xlsx")
    start_idx = max(0, args.start - 1)
    df_subset = df_questions.iloc[start_idx : start_idx + args.limit]

    print(f"Toplam {len(df_subset)} soru için A/B testi başlıyor...")
    print(f"  1. Model: {MODEL_1_ID} (temp={MODEL_1_TEMPERATURE})")
    print(f"  2. Model: {MODEL_2_ID} (temp={MODEL_2_TEMPERATURE})")
    print(f"  3. Model: {MODEL_3_ID} (temp={MODEL_3_TEMPERATURE})")
    print(f"  Hakem   : {JUDGE_MODEL_ID} (temp={JUDGE_TEMPERATURE})\n", flush=True)

    checkpoint_file = f"{args.prefix}checkpoint_results.json"
    results_map = {}

    rows_to_process = [row.to_dict() for _, row in df_subset.iterrows()]

    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_q = {executor.submit(process_single_question, r): int(r["Soru No"]) for r in rows_to_process}
        for future in as_completed(future_to_q):
            q_no = future_to_q[future]
            res = future.result()
            results_map[q_no] = res
            print(
                f"[Tamamlandı {len(results_map)}/{len(df_subset)}] Soru {q_no}: "
                f"{MODEL_1_SHORT_NAME}=%{res[f'{MODEL_1_SHORT_NAME} Doğruluk (%)']} | "
                f"{MODEL_2_SHORT_NAME}=%{res[f'{MODEL_2_SHORT_NAME} Doğruluk (%)']} | "
                f"{MODEL_3_SHORT_NAME}=%{res[f'{MODEL_3_SHORT_NAME} Doğruluk (%)']}",
                flush=True
            )
            with open(checkpoint_file, "w", encoding="utf-8") as f:
                json.dump(list(results_map.values()), f, ensure_ascii=False, indent=2)

    results = [results_map[int(row["Soru No"])] for _, row in df_subset.iterrows()]
    df_results = pd.DataFrame(results)

    # 1. Detaylı Tablo
    detailed_file = f"{args.prefix}model_cevaplari_ve_dogruluk.xlsx"
    try:
        df_results.to_excel(detailed_file, index=False)
        format_detailed_excel(detailed_file)
    except PermissionError:
        detailed_file = f"{args.prefix}model_cevaplari_ve_dogruluk_yeni.xlsx"
        df_results.to_excel(detailed_file, index=False)
        format_detailed_excel(detailed_file)

    # 2. Genel Toplam Doğruluk Tablosu (Her model ayrı ayrı)
    scores_1 = df_results[f"{MODEL_1_SHORT_NAME} Doğruluk (%)"]
    scores_2 = df_results[f"{MODEL_2_SHORT_NAME} Doğruluk (%)"]
    scores_3 = df_results[f"{MODEL_3_SHORT_NAME} Doğruluk (%)"]

    summary_rows = [
        {
            "Model Sıra": 1,
            "Model Kısa Adı": MODEL_1_SHORT_NAME,
            "OpenRouter Model ID": MODEL_1_ID,
            "Temperature": MODEL_1_TEMPERATURE,
            "Test Edilen Soru Sayısı": len(df_results),
            "Tam Doğru (%100) Soru Sayısı": int((scores_1 == 100).sum()),
            "Genel Toplam Doğruluk Oranı (%)": f"%{round(scores_1.mean(), 2)}"
        },
        {
            "Model Sıra": 2,
            "Model Kısa Adı": MODEL_2_SHORT_NAME,
            "OpenRouter Model ID": MODEL_2_ID,
            "Temperature": MODEL_2_TEMPERATURE,
            "Test Edilen Soru Sayısı": len(df_results),
            "Tam Doğru (%100) Soru Sayısı": int((scores_2 == 100).sum()),
            "Genel Toplam Doğruluk Oranı (%)": f"%{round(scores_2.mean(), 2)}"
        },
        {
            "Model Sıra": 3,
            "Model Kısa Adı": MODEL_3_SHORT_NAME,
            "OpenRouter Model ID": MODEL_3_ID,
            "Temperature": MODEL_3_TEMPERATURE,
            "Test Edilen Soru Sayısı": len(df_results),
            "Tam Doğru (%100) Soru Sayısı": int((scores_3 == 100).sum()),
            "Genel Toplam Doğruluk Oranı (%)": f"%{round(scores_3.mean(), 2)}"
        }
    ]

    df_summary = pd.DataFrame(summary_rows)
    summary_file = f"{args.prefix}genel_toplam_dogruluk.xlsx"
    try:
        df_summary.to_excel(summary_file, index=False)
        format_summary_excel(summary_file)
    except PermissionError:
        summary_file = f"{args.prefix}genel_toplam_dogruluk_yeni.xlsx"
        df_summary.to_excel(summary_file, index=False)
        format_summary_excel(summary_file)

    print(f"\nTamamlandı!\n1) Detaylı Tablo: {detailed_file}\n2) Genel Özet Tablosu: {summary_file}")

if __name__ == "__main__":
    main()
