import csv

with open("data/callcenter_hourly_report.csv", encoding="utf-8") as f:
    reader = csv.DictReader(f)
    rows = [row for row in reader]

sick_candidates = [
    row for row in rows
    if row["shift"] == "夜勤" and float(row["answer_rate"]) <= 0.70
]

print(f"応答率0.70以下の夜勤レコード：{len(sick_candidates)}件")
for row in sick_candidates[:5]:
    print(row["business_date"], row["hour"], row["answer_rate"])