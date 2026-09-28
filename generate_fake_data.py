"""
generate_fake_data.py
コールセンターの架空データを生成し、csvファイルに保存する。

役割：
- 日付・シフト別の対応件数（callcenter_report.csv）
- 時間帯別（0～23時）の入電数・応答数・応答率・計画入電数と差異
  (callcenter_hourly_report.csv)
を、ランダムな架空の数値として作り出すこと。

SLA目標に対する評価（計画応答数・応答数差異・SLA判定）は
Excel側のSettingsシート（SLA_Target）を参照してVBAで計算する。
"""

import csv
import random
from datetime import date, timedelta

START_DATE = date(2025, 9, 1)
SEED = 20250901
DAYS = 365
CSV_PATH = "data/callcenter_report.csv"
FIELDNAMES = ["date", "shift", "auth_request_count", "cardholder_inquiry_count", "lost_stolen_count"]

# 1日全体の計画入電数（日勤＋夜勤の合計）
DAILY_PLAN = {
    "auth_request_count":       300,
    "cardholder_inquiry_count": 1500,
    "lost_stolen_count":        1000,
}

JITTER_RANGE = (0.90, 1.10)  # ±10%

# ==============================================================
# 時間帯別データ生成
# ==============================================================

HOURLY_CSV_PATH = "data/callcenter_hourly_report.csv"
HOURLY_FIELDNAMES = [
    "date", "business_date", "hour", "shift",
    "auth_request_count", "cardholder_inquiry_count", "lost_stolen_count",
    "call_count", "planned_call_count", "call_variance",
    "answer_rate", "answered_count",
]

HOUR_TO_SHIFT = {}
for h in range(9, 20):
    HOUR_TO_SHIFT[h] = "日勤"
for h in list(range(20, 24)) + list(range(0, 9)):
    HOUR_TO_SHIFT[h] = "夜勤"

SHIFT_ORDER = {"日勤": 0, "夜勤": 1}

BUSINESS_DAY_HOURS = list(range(9, 24)) + list(range(0, 9))

# 回線別の時間帯ウェイト
CATEGORY_HOUR_WEIGHTS = {
    "auth_request_count": {
        # 9時台に立ち上がり、夕方〜20時過ぎにかけて減る。22時以降は極端に少ない
        9: 0.5,
        10: 1.0, 11: 1.0, 12: 1.0, 13: 1.0, 14: 1.0, 15: 1.0, 16: 1.0,
        17: 0.8, 18: 0.6, 19: 0.4,
        20: 0.3, 21: 0.2, 22: 0.1, 23: 0.1,
        0: 0.05, 1: 0.05, 2: 0.05, 3: 0.05, 4: 0.05,
        5: 0.05, 6: 0.05, 7: 0.05, 8: 0.05,
    },
    "cardholder_inquiry_count": {
        # 正規案内9〜17時、夜間はオーソリセンターで受電（日勤の半分以下）
        9: 0.6, 10: 1.0, 11: 1.1, 12: 0.8, 13: 1.0, 14: 1.1,
        15: 1.0, 16: 0.9, 17: 0.8, 18: 0.6, 19: 0.4,
        20: 0.4, 21: 0.3, 22: 0.3, 23: 0.2,
        0: 0.2, 1: 0.1, 2: 0.1, 3: 0.1, 4: 0.1,
        5: 0.1, 6: 0.2, 7: 0.3, 8: 0.4,
    },
    "lost_stolen_count": {
        # 日中ほぼ横ばい、夜にかけて減る山型、日勤が多い
        9: 0.8, 10: 1.0, 11: 1.0, 12: 1.0, 13: 1.0, 14: 1.0,
        15: 0.9, 16: 0.8, 17: 0.7, 18: 0.6, 19: 0.5,
        20: 0.6, 21: 0.5, 22: 0.4, 23: 0.3,
        0: 0.3, 1: 0.2, 2: 0.2, 3: 0.2, 4: 0.2,
        5: 0.3, 6: 0.4, 7: 0.5, 8: 0.6,
    },
}

# 回線ごとの重み合計を事前計算
CATEGORY_WEIGHT_TOTAL = {
    category: sum(weights.values())
    for category, weights in CATEGORY_HOUR_WEIGHTS.items()
}


def _planned_for_hour(hour: int) -> dict:
    """その時間帯の計画入電数を回線別に返す"""
    return {
        category: DAILY_PLAN[category] * CATEGORY_HOUR_WEIGHTS[category][hour] / CATEGORY_WEIGHT_TOTAL[category]
        for category in DAILY_PLAN
    }


def _normal_answer_rate() -> float:
    """平常時の応答率（乱数の生値）。5%の確率で0.85〜0.90に落ちる"""
    if random.random() < 0.05:
        return random.uniform(0.85, 0.90)
    return random.uniform(0.90, 0.98)


def generate_one_hour(target_date: date, business_date: date, hour: int) -> dict:
    shift = HOUR_TO_SHIFT[hour]
    planned_by_category = _planned_for_hour(hour)
    planned_call_count = round(sum(planned_by_category.values()))

    actual_counts = {
        category: max(0, round(base * random.uniform(*JITTER_RANGE)))
        for category, base in planned_by_category.items()
    }
    call_count = sum(actual_counts.values())
    call_variance = call_count - planned_call_count

    # 応答率の生値（工程⑥以降は「平常値と発生中の事象の最小値」をここで決める）
    raw_rate = _normal_answer_rate()

    # 応答数を確定させてから、応答率を実数から計算し直す（CSV内の矛盾を防ぐ）
    answered_count = round(call_count * raw_rate)
    answer_rate = round(answered_count / call_count, 3) if call_count > 0 else 0.0

    return {
        "date": target_date.isoformat(),
        "business_date": business_date.isoformat(),
        "hour": hour,
        "shift": shift,
        "auth_request_count": actual_counts["auth_request_count"],
        "cardholder_inquiry_count": actual_counts["cardholder_inquiry_count"],
        "lost_stolen_count": actual_counts["lost_stolen_count"],
        "call_count": call_count,
        "planned_call_count": planned_call_count,
        "call_variance": call_variance,
        "answer_rate": answer_rate,
        "answered_count": answered_count,
    }


def generate_all_hours() -> list:
    records = []
    for i in range(DAYS):
        business_date = START_DATE + timedelta(days=i)
        for hour in BUSINESS_DAY_HOURS:
            if hour >= 9:
                target_date = business_date
            else:
                target_date = business_date + timedelta(days=1)
            records.append(generate_one_hour(target_date, business_date, hour))
    return records


def save_hourly_csv(records: list) -> None:
    with open(HOURLY_CSV_PATH, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=HOURLY_FIELDNAMES)
        writer.writeheader()
        writer.writerows(records)


# ==============================================================
# 日次CSV（時間帯別を業務日付×シフトで合計して生成）
# ==============================================================

def aggregate_to_daily(hourly_records: list) -> list:
    """時間帯別レコードを業務日付×シフトで合計し、日次レコードを返す"""
    totals = {}
    for rec in hourly_records:
        key = (rec["business_date"], rec["shift"])
        if key not in totals:
            totals[key] = {
                "date": rec["business_date"],
                "shift": rec["shift"],
                "auth_request_count": 0,
                "cardholder_inquiry_count": 0,
                "lost_stolen_count": 0,
            }
        totals[key]["auth_request_count"]       += rec["auth_request_count"]
        totals[key]["cardholder_inquiry_count"] += rec["cardholder_inquiry_count"]
        totals[key]["lost_stolen_count"]        += rec["lost_stolen_count"]

    # 業務日付→シフト（日勤→夜勤の時系列順）でソート
    return sorted(totals.values(), key=lambda r: (r["date"], SHIFT_ORDER[r["shift"]]))


def save_to_csv(records: list) -> None:
    with open(CSV_PATH, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(records)


if __name__ == "__main__":
    random.seed(SEED)

    # 時間帯別を先に生成
    hourly_records = generate_all_hours()
    save_hourly_csv(hourly_records)
    print(f"{len(hourly_records)}件のデータ({DAYS}日分 x 24時間)を{HOURLY_CSV_PATH}に保存しました。")

    # 日次は時間帯別を合計して生成（数字の一貫性を保つ）
    daily_records = aggregate_to_daily(hourly_records)
    save_to_csv(daily_records)
    print(f"{len(daily_records)}件のデータ({DAYS}日分 x 2シフト)を{CSV_PATH}に保存しました。")