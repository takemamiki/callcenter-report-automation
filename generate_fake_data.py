"""
generate_fake_data.py
コールセンターの架空データを生成し、csvファイルに保存する。

役割：
- 日付・シフト別の対応件数（従来のcallcenter_report.csv）
- 時間帯別（0～23時）の着信数・応答数・応答率・計画値と差異
(callcenter_hourly_report.csv)
を、ランダムな架空の数値として作り出すこと。
"""

import csv
import random
from datetime import date, timedelta

START_DATE = date(2025, 9, 1)
DAYS = 365
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


# ==============================================================
# ここから時間帯別（0～23時）データ生成
# ==============================================================

HOURLY_CSV_PATH = "data/callcenter_hourly_report.csv"
HOURLY_FIELDNAMES = [
    "date", "hour", "shift",
    "auth_request_count", "cardholder_inquiry_count", "lost_stolen_count",
    "call_count", "planned_call_count", "call_variance",
    "answer_rate", "answered_count", "planned_answered_count",
"answered_variance",
]

TARGET_ANSWER_RATE = 0.95
JITTER_RANGE = (0.8, 1.2)

HOUR_TO_SHIFT = {}
for h in range(9, 20):
    HOUR_TO_SHIFT[h] = "日勤"
for h in list(range(20, 24)) + list(range(0, 9)):
    HOUR_TO_SHIFT[h] = "夜勤"

DAY_HOUR_WEIGHTS = {
    9: 0.6, 10: 1.0, 11: 1.1, 12: 0.8, 13: 1.0, 14: 1.1,
    15: 1.0, 16: 0.9, 17: 0.8, 18: 0.6, 19: 0.4,    
}
NIGHT_HOUR_WEIGHTS = {
    20: 1.2, 21: 1.1, 22: 1.0, 23: 0.9, 0: 0.8, 1: 0.6,
    2: 0.4, 3: 0.3, 4: 0.3, 5: 0.4, 6: 0.5, 7: 0.7, 8: 0.9,
}
HOUR_WEIGHTS = {**DAY_HOUR_WEIGHTS, **NIGHT_HOUR_WEIGHTS}

DAY_WEIGHT_TOTAL = sum(DAY_HOUR_WEIGHTS.values())
NIGHT_WEIGHT_TOTAL = sum(NIGHT_HOUR_WEIGHTS.values())
SHIFT_WEIGHT_TOTAL = {"日勤": DAY_WEIGHT_TOTAL, "夜勤": NIGHT_WEIGHT_TOTAL}

SHIFT_DAILY_MIDPOINT = {
    shift: {
        category: sum(range_) / 2
        for category, range_ in categories.items()
    }
    for shift, categories in SHIFT_RANGES.items()
}


def _hour_ratio(hour: int) -> float:
    shift = HOUR_TO_SHIFT[hour]
    return HOUR_WEIGHTS[hour] / SHIFT_WEIGHT_TOTAL[shift]


def generate_one_hour(target_date: date, hour: int) -> dict:
    shift = HOUR_TO_SHIFT[hour]
    ratio = _hour_ratio(hour)
    midpoints = SHIFT_DAILY_MIDPOINT[shift]

    planned_by_category = {
        category: mid * ratio for category, mid in midpoints.items()
    }
    planned_call_count = round(sum(planned_by_category.values()))

    actual_counts = {
        category: max(0, round(base * random.uniform(*JITTER_RANGE)))
        for category, base in planned_by_category.items()
    }
    call_count = sum(actual_counts.values())
    call_variance = call_count - planned_call_count

    answer_rate = round(random.uniform(0.90, 0.98), 3)
    answered_count = round(call_count * answer_rate)
    planned_answered_count = round(planned_call_count * TARGET_ANSWER_RATE)
    answered_variance = answered_count - planned_answered_count

    record = {
        "date": target_date.isoformat(),
        "hour": hour,
        "shift":shift,
        "auth_request_count": actual_counts["auth_request_count"],
        "cardholder_inquiry_count": actual_counts["cardholder_inquiry_count"],
        "lost_stolen_count": actual_counts["lost_stolen_count"],
        "call_count": call_count,
        "planned_call_count": planned_call_count,
        "call_variance": call_variance,
        "answer_rate": answer_rate,
        "answered_count": answered_count,
        "planned_answered_count": planned_answered_count,
        "answered_variance": answered_variance,
    }
    return record


def generate_all_hours() -> list:
    records = []
    for i in range(DAYS):
        target_date = START_DATE + timedelta(days=i)
        for hour in range(24):
            records.append(generate_one_hour(target_date, hour))
    return records


def save_hourly_csv(records: list) -> None:
    with open(HOURLY_CSV_PATH, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=HOURLY_FIELDNAMES)
        writer.writeheader()
        writer.writerows(records)


if __name__ =="__main__":
    records = generate_all_days()
    save_to_csv(records)
    print(f"{len(records)}件のデータ({DAYS}日分 x 2シフト)を{CSV_PATH}に保存しました。  ")

    hourly_records = generate_all_hours()
    save_hourly_csv(hourly_records)
    print(f"{len(hourly_records)}件のデータ({DAYS}日分 x 24時間)を{HOURLY_CSV_PATH}に保存しました。")