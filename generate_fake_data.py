"""
generate_fake_data.py
夜間コールセンターの架空データを生成し、csvファイルに保存する。
役割：日付・シフト・対応件数・エラー件数を、ランダムな架空の数値として作り出すこと。
"""
import csv
import random
from datetime import date, timedelta

START_DATE = date(2026, 8, 1)
DAYS = 30 
SHIFT = "夜間"
CSV_PATH = "data/callcenter_report.csv"
FIELDNAMES = ["date", "shift", "auth_request_count", "cardholder_inquiry_count", "lost_stolen_count"]

def generate_one_day(target_date: date) -> dict:
    auth_request_count = random.randint(50, 120)
    cardholder_inquiry_count = random.randint(20, 60)
    lost_stolen_count = random.randint(50, 80)
    record = {
        "date": target_date.isoformat(),
        "shift": SHIFT,
        "auth_request_count": auth_request_count,
        "cardholder_inquiry_count": cardholder_inquiry_count,
        "lost_stolen_count": lost_stolen_count,
    }
    return record

def generate_all_days() -> list:
    records = []
    for i in range(DAYS):
        target_date = START_DATE + timedelta(days=i)
        record = generate_one_day(target_date)
        records.append(record)
    return records

def save_to_csv(records: list) -> None:
    with open(CSV_PATH, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(records)

if __name__ =="__main__":
    records = generate_all_days()
    save_to_csv(records)
    print(f"{len(records)}日分のデータを{CSV_PATH}に保存しました。  ")

