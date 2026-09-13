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
CSV_PATH = "data/callcenter_report.csv"
FIELDNAMES = ["date", "shift", "auth_request_count", "cardholder_inquiry_count", "lost_stolen_count"]

# シフトごとの乱数レンジ（仮の値。実際の実務感覚に合わせて後で調整予定。）
SHIFT_RANGES = {
    "日勤": {
        "auth_request_count": (70, 120),
        "cardholder_inquiry_count": (30, 60),
        "lost_stolen_count": (30, 60),
    },
    "夜勤": {
        "auth_request_count": (20, 60),
        "cardholder_inquiry_count": (10, 30),
        "lost_stolen_count": (50, 100),
    },
}

def generate_one_shift(target_date: date, shift: str) -> dict:
    ranges = SHIFT_RANGES[shift]
    record = {
        "date": target_date.isoformat(),
        "shift": shift,
        "auth_request_count": random.randint(*ranges["auth_request_count"]),
        "cardholder_inquiry_count": random.randint(*ranges["cardholder_inquiry_count"]),
        "lost_stolen_count": random.randint(*ranges["lost_stolen_count"]),
    }
    return record

def generate_all_days() -> list:
    records = []
    for i in range(DAYS):
        target_date = START_DATE + timedelta(days=i)
        for shift in SHIFT_RANGES:
            record = generate_one_shift(target_date, shift)
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
    print(f"{len(records)}件のデータ({DAYS}日分 x 2シフト)を{CSV_PATH}に保存しました。  ")

